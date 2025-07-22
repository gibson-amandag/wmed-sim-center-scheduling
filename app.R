library(shiny)
library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)
library(zip)
library(colourpicker)
library(shinyTime)
library(shinyjs)
library(lubridate)
library(hms)

fraction_to_posix <- function(frac) {
  if (is.na(frac) || frac == "") return(NULL)
  h <- floor(frac * 24)
  m <- round((frac * 24 - h) * 60)
  if (m == 60) {
    h <- h + 1
    m <- 0
  }
  strptime(sprintf("2000-01-01 %02d:%02d", h, m), "%Y-%m-%d %H:%M")
}

# Helper to load all sheets
load_data <- function(file) {
  required_sheets <- c("studentInfo", "groupInfo", "fillColor", "timeBlockInfo", "schedule", "faculty")
  sheets <- openxlsx::getSheetNames(file)
  missing <- setdiff(required_sheets, sheets)
  if (length(missing) > 0) {
    showNotification(
      paste0(
        "The following required sheets are missing from the uploaded file: ",
        paste(missing, collapse = ", ")
      ),
      type = "error",
      duration = 10
    )
    return(NULL)
  }
  list(
    studentInfo   = read.xlsx(file, sheet = "studentInfo", detectDates = TRUE),
    groupInfo     = read.xlsx(file, sheet = "groupInfo", detectDates = TRUE),
    fillColor     = read.xlsx(file, sheet = "fillColor", detectDates = TRUE),
    timeBlockInfo = read.xlsx(file, sheet = "timeBlockInfo", detectDates = TRUE),
    schedule      = read.xlsx(file, sheet = "schedule", detectDates = TRUE),
    faculty       = read.xlsx(file, sheet = "faculty", detectDates = TRUE)
  )
}

# Generate group schedules
generate_group_schedules <- function(data) {
  schedules <- list()

  # Detect faculty assignment mode
  faculty_by_student <- all(c("groupNum", "studentNum", "faculty") %in% names(data$faculty))
  faculty_by_room <- any(grepl("^group[0-9]+$", names(data$faculty)))

  for (group in unique(as.character(data$studentInfo$groupNum))) {
    group_students <- data$studentInfo %>% filter(groupNum == group)
    group_meta <- data$groupInfo %>% filter(groupNum == group)
    if (nrow(group_meta) == 0) next

    sched <- data$schedule
    time_blocks <- grep("^TimeBlock", names(sched), value = TRUE)
    tb_info <- data$timeBlockInfo

    if (faculty_by_room) {
      # --- By room: original logic ---
      # Ensure all group columns are character before pivoting
      data$faculty <- data$faculty %>%
        mutate(across(starts_with("group"), as.character))
      faculty_long <- data$faculty %>%
        pivot_longer(
          cols = starts_with("group"),
          names_to = "groupNum",
          values_to = "faculty",
          names_prefix = "group"
        )
      faculty_long$groupNum <- as.character(faculty_long$groupNum)
      sched_with_faculty <- sched %>%
        left_join(
          faculty_long %>% filter(groupNum == as.character(group)),
          by = c("shortKey")
        ) %>%
        mutate(
          faculty = ifelse(!is.na(faculty), faculty, ifelse(!is.null(sched$faculty), sched$faculty, NA))
        )
    } else {
      # No faculty info
      sched_with_faculty <- sched
      sched_with_faculty$faculty <- NA
    }

    # Wide version: replace studentNum with "studentNum. lastName, firstName"
    wide_sched <- sched_with_faculty
    for (tb in time_blocks) {
      wide_sched[[tb]] <- sapply(seq_len(nrow(wide_sched)), function(i) {
        sn <- sched[[tb]][i]
        if (is.na(sn) || sn == "") {
          return("")
        }
        sn_vec <- unlist(strsplit(as.character(sn), ","))
        sn_vec <- sn_vec[sn_vec != ""]
        labels <- sapply(sn_vec, function(snv) {
          stu <- group_students[group_students$studentNum == as.integer(snv), ]
          fac <- NULL
          if (faculty_by_student) {
            fac_row <- data$faculty[data$faculty$groupNum == group & data$faculty$studentNum == as.integer(snv), ]
            if (nrow(fac_row) > 0 && !is.na(fac_row$faculty[1]) && fac_row$faculty[1] != "") {
              fac <- fac_row$faculty[1]
            }
          }
          label <- if (nrow(stu) > 0) {
            paste0(stu$studentNum, ". ", stu$lastName, ", ", stu$firstName)
          } else {
            as.character(snv)
          }
          if (!is.null(fac)) {
            paste0(label, " (", fac, ")")
          } else {
            label
          }
        })
        paste(labels, collapse = "; ")
      })
    }

    sched_with_faculty[time_blocks] <- lapply(sched_with_faculty[time_blocks], as.character)

    # Long version: one row per station/time block
    long_sched <- tidyr::pivot_longer(
      sched_with_faculty,
      cols = all_of(time_blocks),
      names_to = "timeBlock",
      values_to = "studentNum"
    ) %>%
      # Split comma-separated studentNums into multiple rows
      tidyr::separate_rows(studentNum, sep = ",") %>%
      mutate(studentNum = trimws(studentNum)) %>%
      filter(studentNum != "") %>%
      mutate(studentNum = as.integer(studentNum)) %>%
      left_join(group_students, by = "studentNum") %>%
      left_join(data$fillColor, by = "studentNum") %>%
      mutate(
        studentLabel = ifelse(
          !is.na(lastName),
          ifelse(
            faculty_by_student & !is.na(faculty) & faculty != "",
            paste0(studentNum, ". ", lastName, ", ", firstName, " (", faculty, ")"),
            paste0(studentNum, ". ", lastName, ", ", firstName)
          ),
          as.character(studentNum)
        )
      )

    # If by student, add faculty for each time block
    if (faculty_by_student) {
      long_sched <- long_sched %>%
        left_join(
          data$faculty %>% filter(groupNum == group),
          by = c("groupNum", "studentNum")
        ) %>%
        mutate(faculty = faculty.y) %>%
        select(-faculty.x, -faculty.y)
    }

    # Get group date and start time label
    group_date <- group_meta$date[1]
    group_timeOfDay <- group_meta$timeOfDay[1]

    # Find the row in timeBlockInfo for this group's start time label
    tb_row <- which(tb_info$startTimeLabel == group_timeOfDay)
    if (length(tb_row) == 1) {
      # Get time block times for this start time
      timeblock_times <- setNames(
        lapply(seq_along(time_blocks), function(i) {
          start_col <- paste0("Block", i, "_Start")
          end_col <- paste0("Block", i, "_End")
          start_val <- tb_info[[start_col]][tb_row]
          end_val <- tb_info[[end_col]][tb_row]
          # Helper to format fraction to time string
          format_time <- function(val) {
            if (!is.null(val) && !is.na(val)) {
              h <- floor(val * 24)
              m <- round((val * 24 - h) * 60)
              if (m == 60) {
                h <- h + 1
                m <- 0
              }
              sprintf("%02d:%02d", h, m)
            } else {
              ""
            }
          }
          start_str <- format_time(start_val)
          end_str <- format_time(end_val)
          if (start_str != "" && end_str != "") {
            paste0(start_str, " - ", end_str)
          } else if (start_str != "") {
            start_str
          } else {
            NA
          }
        }),
        time_blocks
      )
      group_startTime <- tb_info$arrivalTime[tb_row]
      group_endTime <- tb_info$leaveTime[tb_row]
    } else {
      timeblock_times <- setNames(rep(NA, length(time_blocks)), time_blocks)
      group_startTime <- NA
      group_endTime <- NA
    }

    schedules[[paste0("Group_", group)]] <- list(
      wide = wide_sched,
      long = long_sched,
      date = group_date,
      startTime = group_startTime,
      endTime = group_endTime,
      timeOfDay = group_timeOfDay,
      timeblock_times = timeblock_times,
      faculty_by_student = faculty_by_student
    )
  }
  return(schedules)
}

get_start_time_label <- function(index, start_time_names) {
  if (is.null(index) || is.na(index) || index == "") return(NA)
  idx <- suppressWarnings(as.integer(index))
  if (!is.na(idx) && idx >= 1 && idx <= length(start_time_names)) {
    return(start_time_names[idx])
  } else {
    return(NA)
  }
}

# UI
ui <- fluidPage(
  useShinyjs(),
  titlePanel("Schedule Generator"),
  helpText("Note: When entering numbers, be careful about scrolling over the input box, as it may change the value."),
  fluidRow(
    column(
      width = 12,
      class = "col-md-4 col-lg-3",
      style = "background-color: #f5f5f5; padding: 10px; border-right: 1px solid #ddd;",
      h4("Event Information"),
      textInput("event_nice_name", "Event Name (for display)", value = "", width = "100%"),
      textInput("event_faculty_contact", "Faculty Contact", value = "", width = "100%"),
      textInput("event_file_name", "File Name (for downloads)", value = "", width = "100%"),
      h2("Step 1:"),
      h3("Option (a)"),
      p("Enter the schedule information within the 'Enter Info' and 'Station Assignments' tabs"),
      h3("Option (b)"),
      p("Upload an existing Excel template and then edit within the Enter Info and Station Assignments tab as desired"),
      fileInput("file", "Upload Template", accept = ".xlsx", width = "100%"),
      # p(
      #   "Be sure to check the uploaded information for any errors, such as incorrect group names, or double assignments",
      #   style = "color: red;"
      # ),
      uiOutput("any_errors_warning"),
      h2("Step 2:"),
      p("Click the button below to load the entered information and generate schedules"),
      p(
        "Be sure to click this button again if you make any changes to the schedule information",
        style = "color: blue;"
      ),
      actionButton("load_info", "Generate Schedules", icon = icon("cogs"), class = "btn-primary", width = "100%"),
      h2("Step 3:"),
      p("View the generated room schedules in the 'Generated Schedules' tab or look at the student schedules in the 'Student Schedule' tab"),
      p("You can also review the entered information in table form in the other tabs"),
      h2("Step 4:"),
      p("Download the generated schedules or individual student schedules"),
      downloadButton("download", "Save Schedules to Excel", class = "btn-success", width = "100%"),
      hr(),
      downloadButton("download_students", "Download Individual Student Schedules", class = "btn-info", width = "100%"),
      hr(),
      downloadButton("download_station_schedules", "Download Station Schedules", class = "btn-warning", width = "100%"),
      
    ),
    column(
      width = 12,
      class = "col-md-8 col-lg-9",
      tabsetPanel(
        tabPanel(
          "Enter Info",
          br(),
          tabsetPanel(
            tabPanel(
              "Time Information",
              fluidRow(
                column(12, h3("Start Time Information")),
                column(
                  6, p("How many start times are there in the schedule?"),
                  em("You might have different start times for different groups, e.g. AM/PM")
                ),
                column(6, numericInput("tmpl_num_starttimes", "# of start times", 2, min = 1))
              ),
              uiOutput("tmpl_starttime_names_ui"),
              fluidRow(
                column(12, h3("Time block information")),
                column(
                  6,
                  p("How many time blocks are there in the schedule?"),
                  em("This is the number of time slots for each station, and it can include breaks")
                ),
                column(6, numericInput("tmpl_num_timeblocks", "# of time blocks", 6, min = 1))
              ),
              uiOutput("tmpl_timeblock_times_ui")
            ),
            tabPanel(
              "Group Information",
              fluidRow(
                column(12, h3("Group information")),
                column(
                  6, 
                  p("How many groups of students are you scheduling?"),
                  em("This is the number of groups of students that will complete the stations on different dates/times")
                ),
                column(
                  6, 
                  numericInput("tmpl_num_groups", "# of groups", 2, min = 1)
                )
              ),
              uiOutput("tmpl_group_info_ui"),
              fluidRow(
                column(6, p("What is the maximum number of students per group?")),
                column(6, numericInput("tmpl_max_students", "Max # of students/group", 6, min = 1))
              ),
              uiOutput("tmpl_student_colors_ui")
            ),
            tabPanel(
              "Student Information",
              fluidRow(
                column(12, h3("Total number of students")),
                column(6, p("What is the total number of students across all groups?")),
                column(6, numericInput("tmpl_total_students", "Total # of students", 12, min = 1))
              ),
              fluidRow(
                column(
                  12,
                  uiOutput("tmpl_student_overflow_warning")
                )
              ),
              fluidRow(
                column(12, h3("Student Information")),
                column(12, p("Enter student info below, or click the button to paste from Excel (columns: Last Name, First Name, Group #, Student #).")),
                column(6, actionButton("tmpl_paste_students", "Paste from Excel")),
                column(6, actionButton("tmpl_fix_group_student_num_btn", "(Re)calculate group/student numbers")),
                column(12, 
                  uiOutput("tmpl_student_warning_ui")
                )
              ),
              fluidRow(
                column(
                  12,
                  helpText("Double-click on a cell in the table below to edit it.")
                ),
                column(
                  12,
                  DT::DTOutput("tmpl_student_table")
                )
              )
            ),
            tabPanel(
              "Station Information",
              fluidRow(
                column(12, h3("Station information")),
                column(6, p("How many stations are there in the schedule?")),
                column(6, numericInput("tmpl_num_stations", "# of stations", 6, min = 1))
              ),
              fluidRow(
                column(12,
                  p("Note that 'short key' is used to match stations within the code. These must all be unique. You likely don't need to change these"),
                  uiOutput("tmpl_station_info_ui")
                )
              )
            ),
            tabPanel(
              "Faculty Assignment",
              fluidRow(
                column(12, h3("Faculty Assignment Mode")),
                column(
                  6,
                  radioButtons(
                    "faculty_assign_mode",
                    "Assign faculty by:",
                    choices = c("Room" = "room", "Student" = "student"),
                    selected = "room",
                    inline = TRUE
                  )
                )
              ),
              fluidRow(
                column(
                  12, h3("Faculty Assignments")
                )
              ),
              uiOutput("faculty_assignment_ui")
            )
          )
        ),
        tabPanel(
          "Station Assignments",
          fluidRow(
            column(
              12,
              h3("Assign students to stations"),
              p("For stations that are longer than one time block, assign the same student number back-to-back."),
              p("Leave a station blank if there's a break"),
              actionButton("tmpl_clear_assignments", "Clear All Assignments", icon = icon("eraser"), class = "btn-warning"),
              uiOutput("tmpl_schedule_warning_ui"),
              uiOutput("tmpl_schedule_ui")
            )
          )
        ),
        tabPanel(
          "Generated Schedules",
          uiOutput("scheduleTabs")
        ),
        tabPanel(
          "Student Schedule",
          selectInput("student_select", "Select Student", choices = NULL),
          uiOutput("student_schedule_table")
        ),
        tabPanel(
          "Station Schedule",
          selectInput("station_group_select", "Select Group", choices = NULL),
          selectInput("station_select", "Select Station", choices = NULL),
          uiOutput("station_schedule_table")
        ),
        tabPanel(
          "Event Overview Calendar",
          uiOutput("event_overview_calendar")
        ),
        tabPanel(
          "Review data",
          tabsetPanel(
            tabPanel("Schedule Template", uiOutput("schedule")),
            tabPanel("Student Info", tableOutput("studentInfo")),
            tabPanel("Group Info", tableOutput("groupInfo")),
            tabPanel("Time Blocks", tableOutput("timeBlockInfo")),
            tabPanel("Faculty Info", tableOutput("facultyInfo"))
          )
        )
      )
    )
  )
)

# Server
server <- function(input, output, session) {
  output$raw_schedule_table <- renderTable({
    req(data$schedule)
    data$schedule
  })

  data <- reactiveValues()

  #########################
  ## Template Inputs
  #########################
  # --- Start times ---
  tmpl_inputs <- reactiveValues(
    starttime_names = list(),
    timeblock_times = list(),
    arrival_times = list(),
    end_times = list()
  )

  # --- Group Info ---
  tmpl_group_info <- reactiveValues(groups = list())

  # --- Fill Color ---
  tmpl_fillColor <- reactiveValues(colors = list())

  # ---  Student info ---------
  tmpl_students <- reactiveVal(NULL)

  # --- Station Info ---
  tmpl_station_info <- reactiveValues(stations = list())

  # --- Faculty Assignments ---
  faculty_assignments <- reactiveValues(
    by_room = list(),   # by_room[[group]][[station]] = faculty name
    by_student = list() # by_student[[group]][[studentNum]] = faculty name
  )

  uploadedTables <- reactiveValues()

  anyErrors <- reactiveValues(
    duplicateStudentNums = FALSE,
    studentWarnings = FALSE,
    duplicateStations = FALSE
  )

  ##############################
  ## Observers for template inputs
  ##############################

  # --- Observe and update start time labels ---
  update_tmpl_starttime_names <- function() {
    req(input$tmpl_num_starttimes)
    n <- input$tmpl_num_starttimes
    for (i in seq_len(n)) {
      key <- paste0("tmpl_starttime_name_", i)
      val <- input[[key]]
      if (!is.null(val)) tmpl_inputs$starttime_names[[key]] <- val

      # Also store arrival and end times for each start time
      arrival_key <- paste0("tmpl_arrival_", i)
      end_key <- paste0("tmpl_end_", i)
      arrival <- input[[arrival_key]]
      end <- input[[end_key]]
      if (!is.null(arrival)) tmpl_inputs$arrival_times[[arrival_key]] <- arrival
      if (!is.null(end)) tmpl_inputs$end_times[[end_key]] <- end
    }
    # Remove any extra labels if n decreased
    to_remove <- setdiff(names(tmpl_inputs$starttime_names), paste0("tmpl_starttime_name_", seq_len(n)))
    tmpl_inputs$starttime_names[to_remove] <- NULL

    # Remove any extra arrival and end times if n decreased
    arrival_to_remove <- setdiff(names(tmpl_inputs$arrival_times), paste0("tmpl_arrival_", seq_len(n)))
    tmpl_inputs$arrival_times[arrival_to_remove] <- NULL
    end_to_remove <- setdiff(names(tmpl_inputs$end_times), paste0("tmpl_end_", seq_len(n)))
    tmpl_inputs$end_times[end_to_remove] <- NULL
  }
  observe({
    update_tmpl_starttime_names()
  })

  # --- Observe and update time block times ---
  update_tmpl_timeblock_times <- function (){
    req(input$tmpl_num_timeblocks, input$tmpl_num_starttimes)
    n_tb <- input$tmpl_num_timeblocks
    n_st <- input$tmpl_num_starttimes
    isolate({
      for (st_idx in seq_len(n_st)) {
        for (tb_idx in seq_len(n_tb)) {
          start_key <- paste0("tmpl_timeblock_", st_idx, "_", tb_idx, "_start")
          end_key <- paste0("tmpl_timeblock_", st_idx, "_", tb_idx, "_end")
          start_val <- input[[start_key]]
          end_val <- input[[end_key]]
          if (!is.null(start_val)) tmpl_inputs$timeblock_times[[start_key]] <- start_val
          if (!is.null(end_val)) tmpl_inputs$timeblock_times[[end_key]] <- end_val
        }
      }
      # Remove any extra if n_tb or n_st decreased
      valid_keys <- unlist(lapply(seq_len(n_st), function(st_idx) {
        unlist(lapply(seq_len(n_tb), function(tb_idx) {
          c(
            paste0("tmpl_timeblock_", st_idx, "_", tb_idx, "_start"),
            paste0("tmpl_timeblock_", st_idx, "_", tb_idx, "_end")
          )
        }))
      }))
      to_remove <- setdiff(names(tmpl_inputs$timeblock_times), valid_keys)
      tmpl_inputs$timeblock_times[to_remove] <- NULL
    })
  }

  observe({
    update_tmpl_timeblock_times()
  })

  # --- Observe and update group info ---
  update_tmpl_group_info <- function() {
    req(input$tmpl_num_groups, input$tmpl_num_starttimes)
    n <- input$tmpl_num_groups
    isolate({
      for (i in seq_len(n)) {
        prefix <- paste0("tmpl_group_", i, "_")
        groupNum <- input[[paste0(prefix, "groupNum")]]
        date <- input[[paste0(prefix, "date")]]
        startTime <- input[[paste0(prefix, "startTime")]]
        endTime <- input[[paste0(prefix, "endTime")]]
        timeOfDay <- input[[paste0(prefix, "timeOfDay")]]
        tmpl_group_info$groups[[i]] <- list(
          groupNum = if (!is.null(groupNum)) groupNum else i,
          date = if (!is.null(date)) date else NULL,
          startTime = if (!is.null(startTime)) startTime else "",
          endTime = if (!is.null(endTime)) endTime else "",
          timeOfDay = if (!is.null(timeOfDay)) timeOfDay else NA,
          groupColor = if (!is.null(input[[paste0(prefix, "groupColor")]])) input[[paste0(prefix, "groupColor")]] else "#e0f7fa"
        )
      }
      # Remove extra if n decreased
      if (length(tmpl_group_info$groups) > n) {
        tmpl_group_info$groups <- tmpl_group_info$groups[seq_len(n)]
      }
    })
  }

  observe({
    req(input$tmpl_num_groups, input$tmpl_num_starttimes)
    n <- input$tmpl_num_groups
    update_tmpl_group_info()
  })

  # --- Observe and update fill colors ---
  update_tmpl_fillColor <- function() {
    req(input$tmpl_max_students)
    n <- input$tmpl_max_students
    isolate({
      for (i in seq_len(n)) {
        key <- paste0("tmpl_student_color_", i)
        val <- input[[key]]
        if (!is.null(val)) tmpl_fillColor$colors[[key]] <- val
      }
      # Remove any extra if n decreased
      to_remove <- setdiff(names(tmpl_fillColor$colors), paste0("tmpl_student_color_", seq_len(n)))
      tmpl_fillColor$colors[to_remove] <- NULL
    })
  }
  
  observe({
    update_tmpl_fillColor()
  })

  # --- Observe and update station info ---
  update_tmpl_station_info <- function() {
    req(input$tmpl_num_stations)
    n <- input$tmpl_num_stations
    isolate({
      for (i in seq_len(n)) {
        prefix <- paste0("tmpl_station_", i, "_")
        tmpl_station_info$stations[[i]] <- list(
          shortKey = input[[paste0(prefix, "shortKey")]],
          niceName = input[[paste0(prefix, "niceName")]],
          timeInMin = input[[paste0(prefix, "timeInMin")]],
          room1 = input[[paste0(prefix, "room1")]],
          room2 = input[[paste0(prefix, "room2")]],
          notes = input[[paste0(prefix, "notes")]],
          stationColor = input[[paste0(prefix, "stationColor")]]
        )
      }
      # Remove extras if n decreased
      if (length(tmpl_station_info$stations) > n) {
        tmpl_station_info$stations <- tmpl_station_info$stations[seq_len(n)]
      }
    })
  }

  # Save station info reactively
  observe({
    update_tmpl_station_info()
  })
  
  # --- Faculty Assignment Update Function ---
  update_faculty_assignments <- function() {
    req(input$tmpl_num_groups, input$tmpl_num_stations, input$tmpl_max_students, input$tmpl_num_starttimes)
    isolate({
      if (input$faculty_assign_mode == "room") {
        for (g in seq_len(input$tmpl_num_groups)) {
          for (i in seq_len(input$tmpl_num_stations)) {
            inputId <- paste0("faculty_room_", g, "_", i)
            val <- input[[inputId]]
            if (!is.null(val)) {
              if (is.null(faculty_assignments$by_room[[as.character(g)]])) faculty_assignments$by_room[[as.character(g)]] <- list()
              faculty_assignments$by_room[[as.character(g)]][[as.character(i)]] <- val
            }
          }
        }
      } else {
        for (g in seq_len(input$tmpl_num_groups)) {
          for (s in seq_len(input$tmpl_max_students)) {
            inputId <- paste0("faculty_student_", g, "_", s)
            val <- input[[inputId]]
            if (!is.null(val)) {
              if (is.null(faculty_assignments$by_student[[as.character(g)]])) faculty_assignments$by_student[[as.character(g)]] <- list()
              faculty_assignments$by_student[[as.character(g)]][[as.character(s)]] <- val
            }
          }
        }
      }
    })
  }
  
  # --- Faculty Assignment Observers ---
  observe({
    update_faculty_assignments()
  })


  ###########################
  ## Build UI elements
  ###########################

  # Start time UI
  output$tmpl_starttime_names_ui <- renderUI({
    req(input$tmpl_num_starttimes)
    n <- input$tmpl_num_starttimes

    isolate({
      # Always use tmpl_inputs, not upload
      starttime_names <- tmpl_inputs$starttime_names
      arrival_times <- tmpl_inputs$arrival_times
      end_times <- tmpl_inputs$end_times
      fluidRow(
        column(12, helpText("Enter times as hh:mm (24-hour format, e.g. 07:30 and 17:15)")),
        lapply(seq_len(n), function(i) {
          key <- paste0("tmpl_starttime_name_", i)
          val <- if (!is.null(starttime_names[[key]])) {
            starttime_names[[key]]
          } else if (i == 1) "AM" else if (i == 2) "PM" else paste0("Start", i)
          # Use stored values if available
          arrival_key <- paste0("tmpl_arrival_", i)
          end_key <- paste0("tmpl_end_", i)
          stored_arrival <- arrival_times[[arrival_key]]
          stored_end <- end_times[[end_key]]
          default_arrival <- if (i == 1) strptime("07:30", "%H:%M") else if (i == 2) strptime("12:30", "%H:%M") else strptime("08:00", "%H:%M")
          default_end <- if (i == 1) strptime("12:15", "%H:%M") else if (i == 2) strptime("17:15", "%H:%M") else strptime("12:00", "%H:%M")
          tagList(
            column(4, textInput(key, paste0("Start time label ", i), value = val)),
            column(
              4,
              timeInput(arrival_key, "Participant arrival time", value = if (!is.null(stored_arrival)) stored_arrival else default_arrival, seconds = FALSE)
            ),
            column(
              4,
              timeInput(end_key, "Participant end time", value = if (!is.null(stored_end)) stored_end else default_end, seconds = FALSE)
            )
          )
        })
      )
    })
  })

  # --- UI for time block times for each start time ---
  output$tmpl_timeblock_times_ui <- renderUI({
    req(input$tmpl_num_timeblocks, input$tmpl_num_starttimes)
    n_tb <- input$tmpl_num_timeblocks
    n_st <- input$tmpl_num_starttimes
    buildTimeblockUI(n_st, n_tb)
  })

  buildTimeblockUI <- function(num_starttimes, num_timeblocks){
    update_tmpl_timeblock_times()

    start_names <- sapply(seq_len(num_starttimes), function(i) {
      key <- paste0("tmpl_starttime_name_", i)
      if (!is.null(tmpl_inputs$starttime_names[[key]])) tmpl_inputs$starttime_names[[key]]
      else if (i == 1) "AM" else if (i == 2) "PM" else paste0("Start", i)
    })
    fluidRow(
      lapply(seq_len(num_starttimes), function(st_idx) {
        column(
          6,
          fluidRow(
            column(
              12,
              h5(
                style = "color: #007bff;",
                paste0(
                  "Time Blocks for ",
                  ifelse(
                    !is.null(start_names[st_idx]) && start_names[st_idx] != "",
                    start_names[st_idx],
                    paste0("Start ", st_idx)
                  )
                )
              )
            )
          ),
          lapply(seq_len(num_timeblocks), function(tb_idx) {
            start_key <- paste0("tmpl_timeblock_", st_idx, "_", tb_idx, "_start")
            end_key <- paste0("tmpl_timeblock_", st_idx, "_", tb_idx, "_end")
            label <- start_names[st_idx]
            stored_start <- tmpl_inputs$timeblock_times[[start_key]]
            stored_end <- tmpl_inputs$timeblock_times[[end_key]]
            if (tolower(label) == "am") {
              start_minutes <- 8 * 60 + (tb_idx - 1) * 30
              end_minutes <- start_minutes + 30
              default_start <- strptime(sprintf("%02d:%02d", start_minutes %/% 60, start_minutes %% 60), "%H:%M")
              default_end <- strptime(sprintf("%02d:%02d", end_minutes %/% 60, end_minutes %% 60), "%H:%M")
            } else if (tolower(label) == "pm") {
              start_minutes <- 13 * 60 + (tb_idx - 1) * 30
              end_minutes <- start_minutes + 30
              default_start <- strptime(sprintf("%02d:%02d", start_minutes %/% 60, start_minutes %% 60), "%H:%M")
              default_end <- strptime(sprintf("%02d:%02d", end_minutes %/% 60, end_minutes %% 60), "%H:%M")
            } else {
              default_start <- NA
              default_end <- NA
            }
            fluidRow(
              column(
                6,
                timeInput(start_key, paste0("Block ", tb_idx, " Start"), value = if (!is.null(stored_start)) stored_start else default_start, seconds = FALSE)
              ),
              column(
                6,
                timeInput(end_key, paste0("Block ", tb_idx, " End"), value = if (!is.null(stored_end)) stored_end else default_end, seconds = FALSE)
              )
            )
          })
        )
      })
    )
  }

  # --- UI for group info entry ---
    output$tmpl_group_info_ui <- renderUI({
    req(input$tmpl_num_groups, input$tmpl_num_starttimes)
    n <- input$tmpl_num_groups
    start_time_names <- sapply(seq_len(input$tmpl_num_starttimes), function(i) {
      key <- paste0("tmpl_starttime_name_", i)
      nm <- tmpl_inputs$starttime_names[[key]]
      if (is.null(nm) || nm == "") paste0("Start", i) else nm
    })
    # Create named vector: values = index, names = label
    time_of_day_choices <- setNames(as.character(seq_along(start_time_names)), start_time_names)
    isolate({
      tagList(
        lapply(seq_len(n), function(i) {
          prefix <- paste0("tmpl_group_", i, "_")
          group <- tmpl_group_info$groups[[i]]
          groupNum_val <- if (!is.null(group) && !is.null(group$groupNum)) group$groupNum else paste0("Group ", i)
          date_val <- if (!is.null(group) && !is.null(group$date)) group$date else NULL
          startTime_val <- if (!is.null(group) && !is.null(group$startTime)) group$startTime else ""
          endTime_val <- if (!is.null(group) && !is.null(group$endTime)) group$endTime else ""
          # Default to index as value
          timeOfDay_val <- if (!is.null(group) && !is.null(group$timeOfDay) && group$timeOfDay %in% as.character(seq_along(start_time_names))) {
            group$timeOfDay
          } else {
            as.character(i)
          }
          # Use existing group color if present, else default
          groupColor_val <- if (!is.null(group) && !is.null(group$groupColor) && group$groupColor != "") group$groupColor else "#e0f7fa"
          fluidRow(
            column(3, textInput(paste0(prefix, "groupNum"), paste0("Group ", i, " Name"), value = groupNum_val)),
            column(3, dateInput(paste0(prefix, "date"), "Date", value = if (!is.null(date_val)) date_val else NULL)),
            column(3, selectInput(paste0(prefix, "timeOfDay"), "Time of Day", choices = time_of_day_choices, selected = timeOfDay_val)),
            column(3, colourpicker::colourInput(paste0(prefix, "groupColor"), "Group Color", value = groupColor_val))
          )
        })
      )
    })
  })

  # --- UI for student color pickers ---
  output$tmpl_student_colors_ui <- renderUI({
      req(input$tmpl_max_students)
      n <- input$tmpl_max_students
  
      # Always use tmpl_fillColor, not upload
      default_colors <- c(
        "#FF7C80", "#FFA365", "#FFFF00", "#AEFF5D", "#A6A200", "#97CBFF",
        "#9797FF", "#FAB3FF", "#CC66FF", "#D4D2D2", "#FFE285", "#B3773B",
        "#85FFDF", "#25C6FF", "#6B9572", "#FF8FDA", "#93AA00", "#593315",
        "#F13A13", "#232C16"
      )
  
      tagList(
        fluidRow(
          column(12, h4("Student Colors (for fillColor sheet)"))
        ),
        fluidRow(
          lapply(seq_len(n), function(i) {
            key <- paste0("tmpl_student_color_", i)
            # Priority: tmpl_fillColor > input > default
            if (!is.null(tmpl_fillColor$colors[[key]])) {
              val <- tmpl_fillColor$colors[[key]]
            } else {
              val <- isolate(input[[key]])
            }
            default_val <- if (!is.null(val) && val != "") val else default_colors[(i - 1) %% length(default_colors) + 1]
            column(
              4,
              tags$div(
                paste("Student", i),
                colourpicker::colourInput(key, NULL, value = default_val, showColour = "both")
              )
            )
          })
        )
      )
  })

  output$tmpl_student_overflow_warning <- renderUI({
    req(input$tmpl_total_students, input$tmpl_num_groups, input$tmpl_max_students)
    max_allowed <- input$tmpl_num_groups * input$tmpl_max_students
    if (input$tmpl_total_students > max_allowed) {
      div(
        style = "color: #b30000; font-weight: bold; margin-bottom: 10px;",
        paste(
          "Warning: The total number of students (", input$tmpl_total_students,
          ") exceeds the maximum allowed (", max_allowed,
          ") for", input$tmpl_num_groups, "groups ×", input$tmpl_max_students, "students per group."
        )
      )
    }
  })

  # Initialize student table when total students or groups change
  observeEvent(
    {
      input$tmpl_total_students
      input$tmpl_num_groups
      input$tmpl_max_students
    },
    {
      req(input$tmpl_total_students, input$tmpl_num_groups, input$tmpl_max_students)
      n <- input$tmpl_total_students
      max_per_group <- input$tmpl_max_students

      # Generate new groupNum and studentNum vectors
      groupNum <- rep(seq_len(ceiling(n / max_per_group)), each = max_per_group, length.out = n)
      studentNum <- rep(seq_len(max_per_group), times = ceiling(n / max_per_group), length.out = n)

      # Get existing data if present
      old_df <- tmpl_students()
      new_df <- data.frame(
        lastName = "",
        firstName = "",
        groupNum = as.character(groupNum),
        studentNum = studentNum,
        stringsAsFactors = FALSE
      )

      # If old data exists, copy over matching rows
      if (!is.null(old_df)) {
        min_rows <- min(nrow(old_df), nrow(new_df))
        new_df[seq_len(min_rows), c("lastName", "firstName", "groupNum", "studentNum")] <- old_df[seq_len(min_rows), c("lastName", "firstName", "groupNum", "studentNum")]
      }

      tmpl_students(new_df)
    },
    # ignoreInit = TRUE
  )

  # Render editable table
  output$tmpl_student_table <- DT::renderDT({
    DT::datatable(
      tmpl_students()[, c("lastName", "firstName", "groupNum", "studentNum")],
      editable = TRUE,
      rownames = FALSE,
      options = list(dom = "t", ordering = TRUE, pageLength = 100)
    )
  })

  # Update table on edit
  observeEvent(input$tmpl_student_table_cell_edit, {
    info <- input$tmpl_student_table_cell_edit
    df <- tmpl_students()
    df[info$row, info$col + 1] <- info$value
    tmpl_students(df)
  })

  # Paste from Excel helper
  observeEvent(input$tmpl_paste_students, {
    showModal(modalDialog(
      title = "Paste Student Info from Excel",
      "Paste rows (Last Name, First Name, Group #, Student #) below, one per line, tab or comma separated. You can include just the first two columns (Last Name, First Name) if you want to auto-generate group and student numbers.",
      textAreaInput("tmpl_paste_text", "Paste Here", rows = 10, width = "100%"),
      footer = tagList(
        modalButton("Cancel"),
        actionButton("tmpl_paste_apply", "Apply")
      )
    ))
  })

  observeEvent(input$tmpl_paste_apply, {
    req(input$tmpl_paste_text)
    lines <- strsplit(input$tmpl_paste_text, "\n")[[1]]
    parsed <- do.call(rbind, lapply(lines, function(line) {
      vals <- strsplit(line, "[\t,]")[[1]]
      vals <- trimws(vals)
      # Pad or trim to 4 columns
      length(vals) <- 4
      vals
    }))
    df <- as.data.frame(parsed, stringsAsFactors = FALSE)
    # Assign column names
    names(df) <- c("lastName", "firstName", "groupNum", "studentNum")
    n <- nrow(df)
    max_per_group <- input$tmpl_max_students
  
    # Fill missing groupNum if needed
    if (all(is.na(df$groupNum)) || all(df$groupNum == "")) {
      df$groupNum <- rep(seq_len(ceiling(n / max_per_group)), each = max_per_group, length.out = n)
    }
  
    # Fill missing studentNum if needed
    if (all(is.na(df$studentNum)) || all(df$studentNum == "")) {
      df$studentNum <- rep(seq_len(max_per_group), times = ceiling(n / max_per_group), length.out = n)
    }
  
    # Ensure correct types
    df$groupNum <- as.character(df$groupNum)
    df$studentNum <- as.integer(df$studentNum)
  
    # Only keep the four columns in the right order
    df <- df[, c("lastName", "firstName", "groupNum", "studentNum")]
  
    updateNumericInput(session, "tmpl_total_students", value = nrow(df))
    tmpl_students(df)
    removeModal()
  })

  check_duplicate_student_numbers <- function(df) {
    # Find duplicated groupNum/studentNum pairs (both directions)
    dups <- duplicated(df[, c("groupNum", "studentNum")]) | duplicated(df[, c("groupNum", "studentNum")], fromLast = TRUE)
    if (any(dups)) {
      dup_rows <- which(dups)
      dup_pairs <- unique(df[dups, c("groupNum", "studentNum")])
      # For each duplicated pair, find all row numbers and names where it occurs
      result <- do.call(rbind, lapply(seq_len(nrow(dup_pairs)), function(i) {
        rows <- which(df$groupNum == dup_pairs$groupNum[i] & df$studentNum == dup_pairs$studentNum[i])
        data.frame(
          groupNum = dup_pairs$groupNum[i],
          studentNum = dup_pairs$studentNum[i],
          row = paste(rows, collapse = ", "),
          names = paste(paste0(df$firstName[rows], " ", df$lastName[rows], " (row ", rows, ")"), collapse = ", ")
        )
      }))
      anyErrors$duplicateStudentNums <- TRUE  # Set error flag if duplicates found
      return(result)
    } else {
      anyErrors$duplicateStudentNums <- FALSE  # Clear error flag if no duplicates
      return(NULL)
    }
  }

  output$tmpl_student_warning_ui <- renderUI({
    # React to group name changes
    req(input$tmpl_num_groups)
    group_names <- sapply(seq_len(input$tmpl_num_groups), function(i) input[[paste0("tmpl_group_", i, "_groupNum")]])
    df <- tmpl_students()
    dups <- check_duplicate_student_numbers(df)
  
    # Check for groupNum not in groupInfo
    groupInfo_names <- NULL
    update_tmpl_group_info()
    if (!is.null(tmpl_group_info$groups) && length(tmpl_group_info$groups) > 0) {
      groupInfo_names <- sapply(tmpl_group_info$groups, function(g) {
        if (!is.null(g$groupNum) && g$groupNum != "") as.character(g$groupNum) else NA
      })
      groupInfo_names <- groupInfo_names[!is.na(groupInfo_names)]
    }
    missing_groups <- NULL
    if (!is.null(df) && !is.null(groupInfo_names)) {
      missing_groups <- setdiff(unique(df$groupNum), groupInfo_names)
      missing_groups <- missing_groups[!is.na(missing_groups) & missing_groups != ""]
    }
  
    warnings <- list()
  
    # Duplicate warning
    if (!is.null(dups)) {
      messages <- lapply(seq_len(nrow(dups)), function(i) {
        name_list <- unlist(strsplit(dups$names[i], ",\\s*"))
        if (length(name_list) > 2) {
          names_fmt <- paste(
            paste(name_list[-length(name_list)], collapse = ", "),
            "and", name_list[length(name_list)]
          )
        } else if (length(name_list) == 2) {
          names_fmt <- paste(name_list[1], "and", name_list[2])
        } else {
          names_fmt <- name_list[1]
        }
        paste0(
          names_fmt,
          " are in group ", dups$groupNum[i],
          " and listed as student ", dups$studentNum[i]
        )
      })
      warnings <- c(warnings, list(
        div(
          style = "color: #b30000; font-weight: bold; margin-bottom: 10px;",
          tagList(
            "Warning: Duplicate group/student number pairs found:",
            tags$ul(
              lapply(messages, tags$li)
            )
          )
        )
      ))
    }
  
    # Missing groupNum warning
    if (length(missing_groups) > 0) {
      warnings <- c(warnings, list(
        div(
          style = "color: #b30000; font-weight: bold; margin-bottom: 10px;",
          paste0(
            "Warning: The following group(s) in the student list do(es) not exist in the group info: ",
            paste(missing_groups, collapse = ", ")
          )
        )
      ))
    }
  
    if (length(warnings) > 0) {
      anyErrors$studentWarnings <- TRUE
      tagList(warnings)
    } else {
      anyErrors$studentWarnings <- FALSE
      NULL
    }
  })

  observeEvent(input$tmpl_fix_group_student_num_btn, {
    req(input$tmpl_total_students, input$tmpl_max_students)
    n <- input$tmpl_total_students
    max_per_group <- input$tmpl_max_students

    groupNum <- rep(seq_len(ceiling(n / max_per_group)), each = max_per_group, length.out = n)
    studentNum <- rep(seq_len(max_per_group), times = ceiling(n / max_per_group), length.out = n)

    old_df <- tmpl_students()
    new_df <- data.frame(
      lastName = "",
      firstName = "",
      groupNum = as.character(groupNum),
      studentNum = studentNum,
      stringsAsFactors = FALSE
    )
    # Retain names for as many rows as possible
    if (!is.null(old_df)) {
      min_rows <- min(nrow(old_df), nrow(new_df))
      new_df[seq_len(min_rows), c("lastName", "firstName")] <- old_df[seq_len(min_rows), c("lastName", "firstName")]
    }
    tmpl_students(new_df)
  })


  output$tmpl_station_info_ui <- renderUI({
    req(input$tmpl_num_stations)
    n <- input$tmpl_num_stations
    tagList(
      fluidRow(
        lapply(seq_len(n), function(i) {
          prefix <- paste0("tmpl_station_", i, "_")
          # --- FIX: Ensure station exists ---
          if (length(tmpl_station_info$stations) < i || is.null(tmpl_station_info$stations[[i]])) {
            tmpl_station_info$stations[[i]] <<- list()
          }
          station <- tmpl_station_info$stations[[i]]
          short_key_val <- if (!is.null(station$shortKey)) station$shortKey else paste0("S", i)
          nice_name_val <- if (!is.null(station$niceName)) station$niceName else paste0("Station ", i)
          time_in_min_val <- if (!is.null(station$timeInMin)) station$timeInMin else ""
          room1_val <- if (!is.null(station$room1)) station$room1 else ""
          room2_val <- if (!is.null(station$room2)) station$room2 else ""
          notes_val <- if (!is.null(station$notes)) station$notes else ""
          station_color_val <- if (!is.null(station$stationColor)) station$stationColor else "#FFFFFF"
          column(
            width = 12, class = "col-md-6",
            tagList(
              fluidRow(
                column(12, h4(paste0("Station ", i)))
              ),
              fluidRow(
                column(3, class = "col-md-6", textInput(paste0(prefix, "shortKey"), "Short Key", value = short_key_val)),
                column(6, class = "col-md-6", textInput(paste0(prefix, "niceName"), "Station Name", value = nice_name_val)),
                column(3, class = "col-md-6", numericInput(paste0(prefix, "timeInMin"), "Duration (min)", value = time_in_min_val, min = 0)),
                column(4, class = "col-md-6", colourpicker::colourInput(paste0(prefix, "stationColor"), "Color", value = station_color_val, showColour = "both")),
                column(4, class = "col-md-6", textInput(paste0(prefix, "room1"), "Main Room", value = room1_val)),
                column(4, class = "col-md-6", textInput(paste0(prefix, "room2"), "Additional Room", value = room2_val))
              ),
              fluidRow(
                column(12, textInput(paste0(prefix, "notes"), "Notes", value = notes_val))
              ),
              tags$hr()
            )
          )
        })
      )
    )
  })

  output$tmpl_schedule_ui <- renderUI({
    req(input$tmpl_num_stations, input$tmpl_num_timeblocks, input$tmpl_max_students)
    n_stations <- input$tmpl_num_stations
    n_blocks <- input$tmpl_num_timeblocks
    max_students <- input$tmpl_max_students
    student_choices <- c("Break" = "", as.character(seq_len(max_students)))
  
    # Get student colors from the color pickers
    student_colors <- sapply(seq_len(max_students), function(i) {
      key <- paste0("tmpl_student_color_", i)
      val <- input[[key]]
      if (is.null(val) || val == "") "#FFFFFF" else val
    })
    names(student_colors) <- as.character(seq_len(max_students))
  
    station_names <- sapply(seq_len(n_stations), function(i) {
      input[[paste0("tmpl_station_", i, "_niceName")]]
    })
  
    header <- tags$tr(
      tags$th("Station"),
      lapply(seq_len(n_blocks), function(j) tags$th(paste("Block", j)))
    )
  
    rows <- lapply(seq_len(n_stations), function(i) {
      station_name <- station_names[i]
      if (is.null(station_name) || station_name == "") station_name <- paste("Station", i)
      tags$tr(
        tags$td(station_name),
        lapply(seq_len(n_blocks), function(j) {
          inputId <- paste0("sched_", i, "_", j)
          # Get the selected value for this selectInput
          selected_val <- input[[inputId]]
          # Determine background color
          # Use first student for color if any selected
          if (!is.null(selected_val) && length(selected_val) > 0 && selected_val[1] != "" && selected_val[1] %in% names(student_colors)) {
            bg_color <- student_colors[[selected_val[1]]]
          } else if (!is.null(selected_val) && length(selected_val) == 1 && selected_val == "") {
            bg_color <- "#e8e8e8"
          } else {
            bg_color <- "#FFFFFF"
          }
          tags$td(
            style = paste0("background-color:", bg_color, ";"),
            selectizeInput(inputId, NULL, choices = student_choices, selected = selected_val, multiple = TRUE, width = "100%")
          )
        })
      )
    })
  
    tags$div(
      style = "overflow-x: auto; overflow-y: auto; border:1px solid #ccc; border-radius:4px; padding:8px; background:#fff;",
      tags$table(
        style = "width:100%; border-collapse:collapse;border:",
        tags$thead(header),
        tags$tbody(rows)
      ),
      tags$style(HTML("
        #tmpl_schedule_ui table tr th, #tmpl_schedule_ui table tr td {
          border: 1px solid #333 !important;
          padding: 4px 8px !important;
          vertical-align: middle;
        }
        #tmpl_schedule_ui table tr th {
          background: #f8f9fa;
          font-weight: bold;
          text-align: center;
        }
        #tmpl_schedule_ui table tr td {
          text-align: center;
        }
      "))
    )
  })

  check_duplicate_station_assignments <- function(n_stations, n_blocks, input) {
    warnings <- list()
    for (j in seq_len(n_blocks)) {
      # Gather all student assignments for this block
      selected <- sapply(seq_len(n_stations), function(i) {
        val <- input[[paste0("sched_", i, "_", j)]]
        # If multiple selected, collapse to comma-separated string
        if (is.null(val)) return("")
        if (length(val) == 0) return("")
        paste(as.character(val), collapse = ",")
      })
      # Flatten: split all comma-separated values into individual student numbers
      selected_flat <- unlist(strsplit(as.character(selected), ","))
      selected_flat <- trimws(selected_flat)
      # Remove blanks/breaks
      selected_flat <- selected_flat[selected_flat != "" & !is.na(selected_flat)]
      # Check for duplicates
      dups <- selected_flat[duplicated(selected_flat)]
      if (length(dups) > 0) {
        dups <- unique(dups)
        warnings[[length(warnings) + 1]] <- paste(
          "Block", j, ": Student(s)", paste(dups, collapse = ", "),
          "assigned to multiple stations."
        )
      }
    }
    warnings
  }

  output$tmpl_schedule_warning_ui <- renderUI({
    req(input$tmpl_num_stations, input$tmpl_num_timeblocks)
    n_stations <- input$tmpl_num_stations
    n_blocks <- input$tmpl_num_timeblocks
    warnings <- check_duplicate_station_assignments(n_stations, n_blocks, input)
    if (length(warnings) > 0) {
      warning <- div(
        style = "color: #b30000; font-weight: bold; margin-bottom: 10px;",
        tagList(
          "Warning: The following students are assigned to multiple stations in the same block:",
          tags$ul(lapply(warnings, tags$li))
        )
      )
      anyErrors$duplicateStations <- TRUE
      warning
    } else {
      anyErrors$duplicateStations <- FALSE
      NULL
    }
  })

  observeEvent(input$tmpl_clear_assignments, {
    req(input$tmpl_num_stations, input$tmpl_num_timeblocks)
    n_stations <- input$tmpl_num_stations
    n_blocks <- input$tmpl_num_timeblocks
    for (i in seq_len(n_stations)) {
      for (j in seq_len(n_blocks)) {
        updateSelectInput(session, paste0("sched_", i, "_", j), selected = "")
      }
    }
  })

  # --- Faculty Assignment UI ---
  output$faculty_assignment_ui <- renderUI({
    req(input$tmpl_num_groups, input$tmpl_num_stations, input$tmpl_max_students)
    num_groups <- input$tmpl_num_groups
    num_stations <- input$tmpl_num_stations
    max_students <- input$tmpl_max_students

    num_starttimes <- input$tmpl_num_starttimes

    # Get start time names from tmpl_inputs
    start_time_names <- sapply(seq_len(num_starttimes), function(i) {
      key <- paste0("tmpl_starttime_name_", i)
      nm <- tmpl_inputs$starttime_names[[key]]
      if (is.null(nm) || nm == "") paste0("Start", i) else nm
    })

    update_tmpl_group_info()
    update_faculty_assignments()
    group_panels <- lapply(seq_len(num_groups), function(g) {
      # Use group name from group info if available
      group_label <- {
        group <- tmpl_group_info$groups[[g]]
        if (!is.null(group) && !is.null(group$groupNum) && group$groupNum != "") {
          paste("Group", group$groupNum)
        } else {
          paste("Group", g)
        }
      }
      group_date <- input[[paste0("tmpl_group_", g, "_date")]]
      group_timeOfDay <- input[[paste0("tmpl_group_", g, "_timeOfDay")]]
      time_label <- get_start_time_label(group_timeOfDay, start_time_names)

      group_heading <- tags$h4(
        group_label,
        if (!is.null(group_date) && !is.na(group_date)) {
          paste0(" (", format(as.Date(group_date), "%A, %B %d, %Y"),
                if (!is.null(time_label) && !is.na(time_label)) paste0(", ", time_label) else "",
                ")")
        }
      )
  
      if (input$faculty_assign_mode == "room") {
        # By room: Table for this group, stations as rows
        station_names <- sapply(seq_len(num_stations), function(i) {
          val <- input[[paste0("tmpl_station_", i, "_niceName")]]
          if (is.null(val) || val == "") paste0("Station ", i) else val
        })
        rows <- lapply(seq_len(num_stations), function(i) {
          inputId <- paste0("faculty_room_", g, "_", i)
          val <- NULL
          if (!is.null(faculty_assignments$by_room[[as.character(g)]])) {
            val <- faculty_assignments$by_room[[as.character(g)]][[as.character(i)]]
          }
          if (is.null(val)) val <- ""
          tags$tr(
            tags$td(station_names[i]),
            tags$td(textInput(inputId, NULL, value = val, width = "100%"))
          )
        })
        table_ui <- tags$table(
          style = "width:100%; margin-bottom: 16px;",
          tags$thead(
            tags$tr(
              tags$th("Station"),
              tags$th("Faculty")
            )
          ),
          tags$tbody(rows)
        )
      } else {
        # By student: Table for this group, student numbers as rows
        rows <- lapply(seq_len(max_students), function(s) {
          inputId <- paste0("faculty_student_", g, "_", s)
          val <- NULL
          if (!is.null(faculty_assignments$by_student[[as.character(g)]])) {
            val <- faculty_assignments$by_student[[as.character(g)]][[as.character(s)]]
          }
          if (is.null(val)) val <- ""
          tags$tr(
            tags$td(s),
            tags$td(textInput(inputId, NULL, value = val, width = "100%"))
          )
        })
        table_ui <- tags$table(
          style = "width:100%; margin-bottom: 16px;",
          tags$thead(
            tags$tr(
              tags$th("Student #"),
              tags$th("Faculty")
            )
          ),
          tags$tbody(rows)
        )
      }
      tags$div(
        style = "margin-bottom: 32px; border: 1px solid #ccc; border-radius: 6px; padding: 12px;",
        group_heading,
        table_ui
      )
    })
  
    tagList(group_panels)
  })

  output$any_errors_warning <- renderUI({
    if (anyErrors$duplicateStudentNums) {
      div(
        style = "color: #b30000; font-weight: bold; margin-bottom: 10px;",
        "Warning: Duplicate group/student number pairs found. Please resolve these before proceeding."
      )
    } else if (anyErrors$studentWarnings) {
      div(
        style = "color: #b30000; font-weight: bold; margin-bottom: 10px;",
        "Warning: There are issues with the student data. Please check the warnings in the student information tab."
      )
    } else if (anyErrors$duplicateStations) {
      div(
        style = "color: #b30000; font-weight: bold; margin-bottom: 10px;",
        "Warning: Some students are assigned to multiple stations in the same block. Please resolve these before proceeding."
      )
    } else {
      NULL
    }
  })
  
  #######################
  ## File upload handling
  #######################

  observeEvent(input$file, {
    req(input$file)
    tables <- load_data(input$file$datapath)

    sheets <- openxlsx::getSheetNames(input$file$datapath)
    if ("Event Info" %in% sheets) {
      event_info <- read.xlsx(input$file$datapath, sheet = "Event Info")
      if ("Field" %in% names(event_info) && "Value" %in% names(event_info)) {
        if ("Event Name" %in% event_info$Field) {
          updateTextInput(session, "event_nice_name", value = event_info$Value[event_info$Field == "Event Name"][1])
        }
        if ("Faculty Contact" %in% event_info$Field) {
          updateTextInput(session, "event_faculty_contact", value = event_info$Value[event_info$Field == "Faculty Contact"][1])
        }
      }
    }

    # Infer nice name from file name (remove extension, replace underscores with spaces)
    fname <- input$file$name
    base <- tools::file_path_sans_ext(basename(fname))
    nice <- gsub("_", " ", base)
    updateTextInput(session, "event_file_name", value = base)

    data$studentInfo <- tables$studentInfo
    data$groupInfo <- tables$groupInfo
    data$fillColor <- tables$fillColor
    data$timeBlockInfo <- tables$timeBlockInfo
    data$schedule <- tables$schedule
    data$faculty <- tables$faculty

    uploadedTables$tables <- tables

    updateUInumbersFromUploadedData()

    delay(
      1000,
      {
        updateTimeInfoFromUploadedData(tables$timeBlockInfo)
        delay(
          100,
          {
            updateGroupInfoFromUploadedData(tables$groupInfo, tables$timeBlockInfo$startTimeLabel)
            updateStudentColorsFromUploadedData(tables$fillColor)
            delay(
              100,
              {
                updateStudentTableFromUploadedData(tables$studentInfo)
                delay(
                  100,
                  {
                    updateStationInfoFromUploadedData(tables$schedule)
                    delay(
                      300,
                      {
                        updateFacultyInfoFromUploadedData(tables$faculty)
                        delay(
                          100, 
                          {
                            updateStationAssignmentsFromUploadedData(tables$schedule)
                          }
                        )
                      }
                    )
                  }
                )
              }
            )
          }
        )
      }
    )
  })

  updateUInumbersFromUploadedData <- function() {
    req(uploadedTables$tables)
    tables <- uploadedTables$tables

    # Update numeric inputs based on uploaded data
    updateNumericInput(session, "tmpl_num_starttimes", value = nrow(tables$timeBlockInfo))
    updateNumericInput(session, "tmpl_num_timeblocks", value = length(grep("^Block[0-9]+_Start$", names(tables$timeBlockInfo))))
    updateNumericInput(session, "tmpl_num_groups", value = nrow(tables$groupInfo))
    updateNumericInput(session, "tmpl_max_students", value = max(as.integer(tables$studentInfo$studentNum), na.rm = TRUE))
    updateNumericInput(session, "tmpl_total_students", value = nrow(tables$studentInfo))
    updateNumericInput(session, "tmpl_num_stations", value = nrow(tables$schedule))

    # Update faculty radio button based on uploaded faculty table columns
    if (!is.null(tables$faculty)) {
      faculty_cols <- names(tables$faculty)
      if (all(c("groupNum", "studentNum", "faculty") %in% faculty_cols)) {
        updateRadioButtons(session, "faculty_assign_mode", selected = "student")
      } else if (any(grepl("^group[0-9]+$", faculty_cols))) {
        updateRadioButtons(session, "faculty_assign_mode", selected = "room")
      }
    }
  }

   updateStudentTableFromUploadedData <- function(studentInfo) {
    # Create a new data frame with the correct columns
    new_students <- data.frame(
      lastName = studentInfo$lastName,
      firstName = studentInfo$firstName,
      groupNum = as.character(studentInfo$groupNum),
      studentNum = as.integer(studentInfo$studentNum),
      stringsAsFactors = FALSE
    )
    tmpl_students(new_students)
  }

  updateGroupInfoFromUploadedData <- function (groupInfo, startTimeLabels) {
    for (i in seq_len(nrow(groupInfo))) {
      # Find the index of the matching start time label
      timeOfDay_idx <- which(startTimeLabels == groupInfo$timeOfDay[i])
      updateTextInput(session, paste0("tmpl_group_", i, "_groupNum"), value = groupInfo$groupNum[i])
      updateDateInput(session, paste0("tmpl_group_", i, "_date"), value = groupInfo$date[i])
      if (length(timeOfDay_idx) == 1) {
        updateSelectInput(session, paste0("tmpl_group_", i, "_timeOfDay"), selected = as.character(timeOfDay_idx))
      }
      if ("groupColor" %in% names(groupInfo)) {
        colourpicker::updateColourInput(session, paste0("tmpl_group_", i, "_groupColor"), value = groupInfo$groupColor[i])
      }
    }
  }

  updateStudentColorsFromUploadedData <- function(fillColor) {
    for (i in seq_len(nrow(fillColor))) {
      colourpicker::updateColourInput(session, paste0("tmpl_student_color_", i), value = fillColor$code[i])
    }
  }

  updateTimeInfoFromUploadedData <- function(timeBlockInfo) {
    req(timeBlockInfo)
    # Check structure: must have startTimeLabel, arrivalTime, leaveTime, and at least one BlockX_Start/BlockX_End
    required_cols <- c("startTimeLabel", "arrivalTime", "leaveTime")
    block_start_cols <- grep("^Block[0-9]+_Start$", names(timeBlockInfo), value = TRUE)
    block_end_cols <- grep("^Block[0-9]+_End$", names(timeBlockInfo), value = TRUE)
    if (!all(required_cols %in% names(timeBlockInfo)) || length(block_start_cols) == 0 || length(block_end_cols) == 0) {
      showNotification(
        "Uploaded timeBlockInfo sheet does not have the required columns (startTimeLabel, arrivalTime, leaveTime, BlockX_Start, BlockX_End). Skipping update.",
        type = "error",
        duration = 10
      )
      return(invisible(NULL))
    }

    num_starttimes <- nrow(timeBlockInfo)
    num_timeblocks <- length(block_start_cols)
    # Update number of start times and time blocks
    updateNumericInput(session, "tmpl_num_starttimes", value = num_starttimes)
    updateNumericInput(session, "tmpl_num_timeblocks", value = num_timeblocks)

    for (i in seq_len(num_starttimes)) {
      # Update start time label
      updateTextInput(session, paste0("tmpl_starttime_name_", i), value = timeBlockInfo$startTimeLabel[i])
      # Update arrival and end times (convert Excel fraction to POSIXct)
      arrival_val <- timeBlockInfo$arrivalTime[i]
      end_val <- timeBlockInfo$leaveTime[i]
      if (!is.na(arrival_val)) {
        arrival_time <- fraction_to_posix(arrival_val)
        updateTimeInput(session, paste0("tmpl_arrival_", i), value = arrival_time)
      }
      if (!is.na(end_val)) {
        end_time <- fraction_to_posix(end_val)
        updateTimeInput(session, paste0("tmpl_end_", i), value = end_time)
      }
      # Update each time block's start and end
      for (tb in seq_len(num_timeblocks)) {
        start_col <- paste0("Block", tb, "_Start")
        end_col <- paste0("Block", tb, "_End")
        if (start_col %in% names(timeBlockInfo)) {
          val <- timeBlockInfo[[start_col]][i]
          if (!is.na(val)) {
            t <- fraction_to_posix(val)
            updateTimeInput(session, paste0("tmpl_timeblock_", i, "_", tb, "_start"), value = t)
          }
        }
        if (end_col %in% names(timeBlockInfo)) {
          val <- timeBlockInfo[[end_col]][i]
          if (!is.na(val)) {
            t <- fraction_to_posix(val)
            updateTimeInput(session, paste0("tmpl_timeblock_", i, "_", tb, "_end"), value = t)
          }
        }
      }
    }
  }

  updateStationInfoFromUploadedData <- function(stationInfo) {
    req(stationInfo)
    n <- nrow(stationInfo)
    for (i in seq_len(n)) {
      prefix <- paste0("tmpl_station_", i, "_")
      updateTextInput(session, paste0(prefix, "shortKey"), value = stationInfo$shortKey[i])
      updateTextInput(session, paste0(prefix, "niceName"), value = stationInfo$niceName[i])
      updateNumericInput(session, paste0(prefix, "timeInMin"), value = stationInfo$timeInMin[i])
      updateTextInput(session, paste0(prefix, "room1"), value = stationInfo$room1[i])
      updateTextInput(session, paste0(prefix, "room2"), value = stationInfo$room2[i])
      updateTextInput(session, paste0(prefix, "notes"), value = stationInfo$notes[i])
      colourpicker::updateColourInput(session, paste0(prefix, "stationColor"), value = stationInfo$stationColor[i])
    }
  }

  updateStationAssignmentsFromUploadedData <- function(schedule) {
    req(schedule)
    # Identify time block columns
    timeblock_cols <- grep("^TimeBlock", names(schedule), value = TRUE)
    n_stations <- nrow(schedule)
    n_blocks <- length(timeblock_cols)
    for (i in seq_len(n_stations)) {
      for (j in seq_len(n_blocks)) {
        inputId <- paste0("sched_", i, "_", j)
        val <- schedule[[timeblock_cols[j]]][i]
        # Fix: handle NA explicitly
        selected <- if (!is.null(val) && !is.na(val) && val != "") {
          vals <- unlist(strsplit(as.character(val), ","))
          vals <- trimws(vals)
          vals[vals != ""]
        } else {
          ""
        }
        updateSelectInput(session, inputId, selected = selected)
      }
    }
  }

  updateFacultyInfoFromUploadedData <- function(facultyInfo) {
    req(facultyInfo)
    # By student: columns groupNum, studentNum, faculty
    if (all(c("groupNum", "studentNum", "faculty") %in% names(facultyInfo))) {
      for (i in seq_len(nrow(facultyInfo))) {
        g <- as.character(facultyInfo$groupNum[i])
        s <- as.integer(facultyInfo$studentNum[i])
        val <- facultyInfo$faculty[i]
        inputId <- paste0("faculty_student_", g, "_", s)
        updateTextInput(session, inputId, value = val)
      }
    # By room: columns shortKey, group1, group2, ...
    } else if ("shortKey" %in% names(facultyInfo) && any(grepl("^group", names(facultyInfo)))) {
      group_cols <- grep("^group", names(facultyInfo), value = TRUE)
      for (row_idx in seq_len(nrow(facultyInfo))) {
        for (group_col in group_cols) {
            # Find the group index for this group_col (e.g., group1 -> 1, group2 -> 2, etc.)
            group_index <- which(group_cols == group_col)
            val <- facultyInfo[[group_col]][row_idx]
            input_id <- paste0("faculty_room_", group_index, "_", row_idx)
            updateTextInput(session, input_id, value = val)
        }
      }
    }
  }

  ########################
  ## Render tables and UI elements
  ########################

  output$studentInfo <- renderTable(
    {
      req(data$studentInfo)
      df <- data$studentInfo
      df$groupNum <- as.character(df$groupNum)
      df$studentNum <- as.integer(df$studentNum)
      df
    },
    striped = TRUE,
    bordered = TRUE
  )

  output$facultyInfo <- renderTable(
    {
      req(data$faculty)
      df <- data$faculty
      df
    },
    striped = TRUE,
    bordered = TRUE
  )

  output$groupInfo <- renderTable(
    {
      req(data$groupInfo)
      df <- data$groupInfo
      df <- df %>%
        mutate(
          date = format(as.Date(date))
        )
      df
    },
    striped = TRUE,
    bordered = TRUE
  )


  fraction_to_time <- function(x) {
    if (is.na(x) || x == "") return("")
    h <- floor(x * 24)
    m <- round((x * 24 - h) * 60)
    if (m == 60) {
      h <- h + 1
      m <- 0
    }
    sprintf("%02d:%02d", h, m)
  }


  output$timeBlockInfo <- renderTable({
    req(data$timeBlockInfo)
    df <- data$timeBlockInfo
    # Find columns that are times (arrival, leave, _Start, _End)
    time_cols <- grep("Time$|_Start$|_End$", names(df), value = TRUE)
    for (col in time_cols) {
      df[[col]] <- sapply(df[[col]], fraction_to_time)
    }
    df
  },
    striped = TRUE,
    bordered = TRUE
  )

  output$schedule <- renderUI({
    req(data$schedule, data$fillColor)
    sched <- data$schedule
    fill <- data$fillColor
  
    # Identify time block columns
    timeblock_cols <- grep("^TimeBlock", names(sched), value = TRUE)
    # Build table header
    header <- tags$tr(
      tags$th("Station"),
      lapply(timeblock_cols, tags$th)
    )
  
    # Build table rows with merged cells for consecutive identical studentNum sets
    rows <- lapply(seq_len(nrow(sched)), function(i) {
      row <- sched[i, ]
      # Compose station info for the first column
      station_info <- tags$div(
        tags$b(row$niceName),
        tags$br(),
        if (!is.null(row$room1) && !is.na(row$room1) && row$room1 != "") {
          paste0("Room: ", row$room1)
        },
        if (!is.null(row$room2) && !is.na(row$room2) && row$room2 != "") {
          list(tags$br(), paste0("Room: ", row$room2))
        },
        tags$br(),
        if (!is.null(row$faculty) && !is.na(row$faculty) && row$faculty != "") {
          paste0("Faculty: ", row$faculty)
        },
        tags$br(),
        if (!is.null(row$notes) && !is.na(row$notes) && row$notes != "") {
          paste0("Notes: ", row$notes)
        }
      )
      # Use stationColor if present
      station_style <- if ("stationColor" %in% names(row) && !is.na(row$stationColor) && row$stationColor != "") {
        paste0("background-color:", row$stationColor, ";")
      } else {
        ""
      }
      cells <- list(tags$td(station_info, style = station_style))
      prev_sn_vec_sorted <- NULL
      prev_label <- NULL
      prev_color <- NULL
      prev_textColor <- NULL
      colspan <- 1
      cell_info <- list()
      for (j in seq_along(timeblock_cols)) {
        tb <- timeblock_cols[j]
        val <- row[[tb]]
        # Parse and sort studentNum list for merging
        sn_vec <- unlist(strsplit(as.character(val), ","))
        sn_vec <- trimws(sn_vec)
        sn_vec <- sn_vec[sn_vec != ""]
        sn_vec_sorted <- sort(sn_vec)
        # Use for coloring (first student), but merge logic uses full sorted list
        studentNum <- if (length(sn_vec_sorted) > 0) suppressWarnings(as.integer(sn_vec_sorted[1])) else NA
        cell_label <- if (is.na(val) || val == "") "Break" else val
        if (is.na(val) || val == "") {
          color <- "#717171"
          textColor <- "white"
        } else if (!is.na(studentNum) && studentNum %in% fill$studentNum) {
          color <- fill$code[fill$studentNum == studentNum]
          textColor <- NULL
        } else {
          color <- "#FFFFFF"
          textColor <- NULL
        }
        # Merge logic: compare full sorted studentNum list
        if (j == 1) {
          prev_sn_vec_sorted <- sn_vec_sorted
          prev_label <- cell_label
          prev_color <- color
          prev_textColor <- if (exists("textColor")) textColor else NULL
          colspan <- 1
        } else if (identical(sn_vec_sorted, prev_sn_vec_sorted) && cell_label != "Break") {
          colspan <- colspan + 1
        } else {
          # Add previous cell
          style_str <- paste0("background-color:", prev_color, ";text-align:center;")
          if (!is.null(prev_textColor)) style_str <- paste0(style_str, "color:", prev_textColor, ";")
          cell_info[[length(cell_info) + 1]] <- tags$td(
            prev_label,
            style = style_str,
            colspan = if (colspan > 1) colspan else NULL
          )
          # Start new cell
          prev_sn_vec_sorted <- sn_vec_sorted
          prev_label <- cell_label
          prev_color <- color
          prev_textColor <- if (exists("textColor")) textColor else NULL
          colspan <- 1
        }
        if (exists("textColor", inherits = FALSE)) rm(textColor, inherits = FALSE)
      }
      # Add last cell
      style_str <- if (!is.null(prev_color)) paste0("background-color:", prev_color, ";text-align:center;") else "text-align:center;"
      if (!is.null(prev_textColor)) style_str <- paste0(style_str, "color:", prev_textColor, ";")
      cell_info[[length(cell_info) + 1]] <- tags$td(
        prev_label,
        style = style_str,
        colspan = if (colspan > 1) colspan else NULL
      )
      do.call(tags$tr, c(cells, cell_info))
    })
  
    tags$table(
      id = "schedule_template_table",
      style = "border-collapse:collapse;width:100%;",
      tags$thead(header),
      tags$tbody(rows)
    ) %>%
      tagAppendChild(
        tags$style(HTML("
          #schedule_template_table tr th, #schedule_template_table tr td {
            border: 1px solid #333 !important;
            padding: 8px 12px !important;
          }
        "))
      )
  })

  ########################
  ## Generate schedules
  ########################

  # Automatically generate schedules when all required data is loaded
  observe({
    req(
      data$studentInfo,
      data$groupInfo,
      data$fillColor,
      data$timeBlockInfo,
      data$schedule
    )
    data$schedules <- generate_group_schedules(data)
  })

  output$scheduleTabs <- renderUI({
    req(data$schedules, data$groupInfo)
    # Use group order from groupInfo if available
    if (!is.null(data$groupInfo) && "groupNum" %in% names(data$groupInfo)) {
      sched_names <- paste0("Group_", as.character(data$groupInfo$groupNum))
      sched_names <- sched_names[sched_names %in% names(data$schedules)]
    } else {
      sched_names <- names(data$schedules)
    }
    tabs <- lapply(sched_names, function(name) {
      # Remove "Group_" prefix and underscore for display
      display_name <- sub("^Group_+", "Group ", name)
      tabPanel(display_name, uiOutput(paste0("sched_", name)))
    })
    do.call(tabsetPanel, tabs)
  })

  observe({
    req(data$schedules, data$fillColor, data$groupInfo)
    # Use group order from groupInfo if available
    if (!is.null(data$groupInfo) && "groupNum" %in% names(data$groupInfo)) {
      sched_names <- paste0("Group_", as.character(data$groupInfo$groupNum))
      sched_names <- sched_names[sched_names %in% names(data$schedules)]
    } else {
      sched_names <- names(data$schedules)
    }
    lapply(sched_names, function(name) {
      output[[paste0("sched_", name)]] <- renderUI({
        scheds <- data$schedules[[name]]
        wide_sched <- scheds$wide
        timeblock_times <- scheds$timeblock_times
        group_date <- scheds$date

        # Identify time block columns
        timeblock_cols <- grep("^TimeBlock", names(wide_sched), value = TRUE)

        # Build table header: Station + time labels (from timeblock_times)
        header <- tags$tr(
          tags$th("Station"),
          lapply(timeblock_cols, function(tb) {
            tags$th(if (!is.null(timeblock_times[[tb]])) timeblock_times[[tb]] else tb)
          })
        )

        # Build table rows with color coding and station info in first column
        rows <- lapply(seq_len(nrow(wide_sched)), function(i) {
          row <- wide_sched[i, ]
          # Compose station info for the first column (like template schedule)
          station_info <- tags$div(
            tags$b(row$niceName),
            tags$br(),
            if (!is.null(row$room1) && !is.na(row$room1) && row$room1 != "") {
              paste0("Room: ", row$room1)
            },
            if (!is.null(row$room2) && !is.na(row$room2) && row$room2 != "") {
              list(tags$br(), paste0("Room: ", row$room2))
            },
            tags$br(),
            if (!is.null(row$faculty) && !is.na(row$faculty) && row$faculty != "") {
              paste0("Faculty: ", row$faculty)
            },
            tags$br(),
            if (!is.null(row$notes) && !is.na(row$notes) && row$notes != "") {
              paste0("Notes: ", row$notes)
            }
          )
          # Use stationColor if present
          station_style <- if ("stationColor" %in% names(row) && !is.na(row$stationColor) && row$stationColor != "") {
            paste0("background-color:", row$stationColor, ";")
          } else {
            ""
          }
          cells <- list(tags$td(station_info, style = station_style))
          prev_studentNum <- NULL
          prev_label <- NULL
          prev_color <- NULL
          prev_textColor <- NULL
          colspan <- 1
          cell_info <- list()
          for (j in seq_along(timeblock_cols)) {
            tb <- timeblock_cols[j]
            val <- row[[tb]]
            # Parse and sort studentNum list for merging
            sn_vec <- unlist(strsplit(as.character(val), ","))
            sn_vec <- trimws(sn_vec)
            sn_vec <- sn_vec[sn_vec != ""]
            sn_vec_sorted <- sort(sn_vec)
            # Use for coloring (first student), but merge logic uses full sorted list
            cell_label <- if (is.na(val) || val == "") "Break" else val
            studentNum <- NA
            matches <- regmatches(val, regexpr("^[0-9]+", val))
            if (length(matches) > 0 && matches != "") {
              studentNum <- as.integer(matches)
            }
            # Only color if student name is present (i.e., val contains ". ")
            if (!is.na(studentNum) && grepl("\\. ", val)) {
              color <- "#FFFFFF"
              textColor <- NULL
              if (studentNum %in% data$fillColor$studentNum) {
                color <- data$fillColor$code[data$fillColor$studentNum == studentNum]
              }
            } else {
              color <- NULL
              textColor <- NULL
            }
            # Merge logic: compare full sorted studentNum list
            if (j == 1) {
              prev_sn_vec_sorted <- sn_vec_sorted
              prev_label <- cell_label
              prev_color <- color
              prev_textColor <- if (exists("textColor")) textColor else NULL
              colspan <- 1
            } else if (identical(sn_vec_sorted, prev_sn_vec_sorted) && cell_label != "Break") {
              colspan <- colspan + 1
            } else {
              # Add previous cell
              style_str <- paste0("background-color:", prev_color, ";text-align:center;")
              if (!is.null(prev_textColor)) style_str <- paste0(style_str, "color:", prev_textColor, ";")
              cell_info[[length(cell_info) + 1]] <- tags$td(
                prev_label,
                style = style_str,
                colspan = if (colspan > 1) colspan else NULL
              )
              # Start new cell
              prev_sn_vec_sorted <- sn_vec_sorted
              prev_label <- cell_label
              prev_color <- color
              prev_textColor <- if (exists("textColor")) textColor else NULL
              colspan <- 1
            }
            if (exists("textColor", inherits = FALSE)) rm(textColor, inherits = FALSE)
          }
          # Add last cell
          style_str <- if (!is.null(prev_color)) paste0("background-color:", prev_color, ";text-align:center;") else "text-align:center;"
          if (!is.null(prev_textColor)) style_str <- paste0(style_str, "color:", prev_textColor, ";")
          cell_info[[length(cell_info) + 1]] <- tags$td(
            prev_label,
            style = style_str,
            colspan = if (colspan > 1) colspan else NULL
          )
          do.call(tags$tr, c(cells, cell_info))
        })

        # Date header above the table
        date_header <- if (!is.null(group_date) && !is.na(group_date)) {
          tags$h4(
            style = "text-align:center;margin-bottom:10px;",
            format(as.Date(group_date), "%A, %B %d, %Y")
          )
        } else {
          NULL
        }

        tagList(
          date_header,
          tags$table(
            id = "group_schedule_table",
            style = "border-collapse:collapse;width:100%;margin:auto;",
            tags$thead(header),
            tags$tbody(rows)
          ) %>%
            tagAppendChild(
              tags$style(HTML("
                #group_schedule_table tr th, #group_schedule_table tr td {
                  border: 1px solid #333 !important;
                  padding: 8px 12px !important;
                }
              "))
            )
        )
      })
    })
  })

  # Update group choices for explorer after schedules are generated
  observeEvent(data$schedules, {
    updateSelectInput(session, "explore_group", choices = names(data$schedules))
  })

  # Update student choices for student schedule tab
  observeEvent(data$studentInfo, {
    req(data$studentInfo)
    # Use all students, not just group 1
    choices <- setNames(
      paste0(data$studentInfo$groupNum, "-", data$studentInfo$studentNum),
      paste0(
        "Group ", data$studentInfo$groupNum, " - ",
        data$studentInfo$studentNum, ". ",
        data$studentInfo$lastName, ", ",
        data$studentInfo$firstName
      )
    )
    updateSelectInput(session, "student_select", choices = choices)
  })

  # Show selected student's schedule (nice display)
  output$student_schedule_table <- renderUI({
    req(data$schedules, input$student_select, data$studentInfo)
    # Parse groupNum and studentNum from selection
    sel <- strsplit(input$student_select, "-", fixed = TRUE)[[1]]
    if (length(sel) != 2) {
      return(NULL)
    }
    groupNum <- as.character(sel[1])
    studentNum <- as.integer(sel[2])
    selected_row <- data$studentInfo[data$studentInfo$groupNum == groupNum & data$studentInfo$studentNum == studentNum, ]
    if (nrow(selected_row) == 0) {
      return(NULL)
    }
    lastName <- selected_row$lastName[1]
    firstName <- selected_row$firstName[1]

    # Find the group schedule
    group_name <- paste0("Group_", groupNum)
    if (!group_name %in% names(data$schedules)) {
      return(NULL)
    }
    sched <- data$schedules[[group_name]]
    group_date <- sched$date
    group_start <- sched$startTime
    group_end <- sched$endTime
    timeblock_times <- sched$timeblock_times

    # Get all time blocks for this group
    long_sched <- sched$long
    student_sched <- long_sched %>%
      filter(
        studentNum == !!studentNum,
        lastName == !!lastName,
        firstName == !!firstName,
        groupNum == !!groupNum
      ) %>%
      arrange(timeBlock)

    # If no schedule, return
    if (nrow(student_sched) == 0) {
      return(tags$div("No schedule found for this student."))
    }

    # Build table rows: one per time block, show time instead of "TimeBlockX"
    rows <- lapply(seq_len(nrow(student_sched)), function(i) {
      row <- student_sched[i, ]
      # Get the time for this time block
      tb_time <- if (!is.null(timeblock_times[[row$timeBlock]])) timeblock_times[[row$timeBlock]] else row$timeBlock
      # Compose station info (like template schedule)
      station_info <- tags$div(
        tags$b(row$niceName),
        tags$br(),
        if (!is.null(row$room1) && !is.na(row$room1) && row$room1 != "") {
          paste0("Room: ", row$room1)
        },
        if (!is.null(row$room2) && !is.na(row$room2) && row$room2 != "") {
          list(tags$br(), paste0("Room: ", row$room2))
        },
        tags$br(),
        if (!is.null(row$faculty) && !is.na(row$faculty) && row$faculty != "") {
          paste0("Faculty: ", row$faculty)
        },
        tags$br(),
        if (!is.null(row$notes) && !is.na(row$notes) && row$notes != "") {
          paste0("Notes: ", row$notes)
        }
      )
      # Use stationColor if present
      station_style <- if ("stationColor" %in% names(row) && !is.na(row$stationColor) && row$stationColor != "") {
        paste0("background-color:", row$stationColor, ";")
      } else {
        ""
      }
      tags$tr(
        tags$td(tb_time),
        tags$td(station_info, style = station_style)
      )
    })

    # Table header
    header <- tags$tr(
      tags$th("Time"),
      tags$th("Station Info")
    )

    tagList(
      tags$div(
        tags$h4(paste(firstName, lastName)),
        tags$p(tags$b("Group #:"), groupNum),
        tags$p(tags$b("Student #:"), studentNum),
        tags$p(tags$b("Name:"), paste(lastName, firstName)),
        tags$p(tags$b("Date:"), format(as.Date(group_date), "%A, %B %d, %Y")),
        tags$p(tags$b("Start time:"), format(strptime(format(as_hms(as.numeric(group_start) * 86400)), "%H:%M:%S"), "%I:%M %p")),
        tags$p(tags$b("End time:"), format(strptime(format(as_hms(as.numeric(group_end) * 86400)), "%H:%M:%S"), "%I:%M %p"))
      ),
      tags$table(
        id = "student_schedule_table",  # <-- Added ID here
        style = "border-collapse:collapse;width:100%;margin:auto;",
        tags$thead(header),
        tags$tbody(rows)
      ) %>%
        tagAppendChild(
          tags$style(HTML("
            #student_schedule_table tr th, #student_schedule_table tr td {
              border: 1px solid #333 !important;
              padding: 8px 12px !important;
            }
          "))
        )
    )
  })

  # Update group choices for station schedule tab
  observeEvent(data$groupInfo, {
    req(data$groupInfo)
    choices <- data$groupInfo$groupNum
    updateSelectInput(session, "station_group_select", choices = choices)
  })

  # Update station choices for station schedule tab
  observeEvent(data$schedule, {
    req(data$schedule)
    sched <- data$schedule
    choices <- setNames(
      sched$shortKey,
      sched$niceName
    )

    updateSelectInput(session, "station_select", choices = choices)
  })

  # Show selected student's schedule (nice display)
  output$station_schedule_table <- renderUI({
    req(data$schedules, input$station_select, input$station_group_select)
  
    stationKey <- as.character(input$station_select)
    niceName <- data$schedule$niceName[data$schedule$shortKey == stationKey]
    groupNum <- as.character(input$station_group_select)
    group_name <- paste0("Group_", groupNum)
    if (!group_name %in% names(data$schedules)) return(NULL)
    sched <- data$schedules[[group_name]]
    group_date <- sched$date
    group_start <- sched$startTime
    group_end <- sched$endTime
  
    # Get station info from wide
    station_wide <- sched$wide %>% filter(shortKey == !!stationKey)
    if (nrow(station_wide) == 0) return(tags$div("No schedule found for this station"))
    room1 <- station_wide$room1
    room2 <- station_wide$room2
    notes <- station_wide$notes
    duration <- station_wide$timeInMin
  
    faculty_by_student <- sched$faculty_by_student
    roomFaculty <- if (!faculty_by_student) station_wide$faculty else NULL
    timeblock_times <- sched$timeblock_times
  
    # Get all time blocks for this group/station
    long_sched <- sched$long %>%
      filter(shortKey == !!stationKey, groupNum == !!groupNum) %>%
      arrange(timeBlock)
  
    if (nrow(long_sched) == 0) return(tags$div("No schedule found for this station"))
  
    # Group by timeBlock, collect all students in each block
    block_df <- long_sched %>%
      group_by(timeBlock) %>%
      summarise(
        studentNums = list(sort(unique(studentNum[!is.na(studentNum)]))),
        studentLabels = list(studentLabel[!is.na(studentLabel) & studentLabel != ""]),
        faculties = list(faculty[!is.na(faculty) & faculty != ""]),
        .groups = "drop"
      )
  
    # For each block, create a merge key (sorted studentNums as string)
    block_df$merge_key <- sapply(block_df$studentNums, function(x) paste(sort(x), collapse = ","))
  
    # Now merge consecutive blocks with the same merge_key
    rows <- list()
    n <- nrow(block_df)
    i <- 1
    while (i <= n) {
      merge_key <- block_df$merge_key[i]
      span <- 1
      while (i + span <= n && block_df$merge_key[i + span] == merge_key) {
        span <- span + 1
      }
      # Compose student info
      studentLabels <- unique(unlist(block_df$studentLabels[i]))
      faculties <- unique(unlist(block_df$faculties[i]))
      if (length(studentLabels) > 0) {
        student_info <- paste0("Student: ", paste(studentLabels, collapse = "; "))
      } else {
        student_info <- "Break"
      }
      if (faculty_by_student && length(faculties) > 0) {
        student_info <- tagList(student_info, tags$br(), paste0("Faculty: ", paste(faculties, collapse = "; ")))
      }
      # Color: use first studentNum if present
      studentNum <- if (length(block_df$studentNums[[i]]) > 0) block_df$studentNums[[i]][1] else NA
      student_color <- if (!is.na(studentNum) && studentNum %in% data$fillColor$studentNum) {
        data$fillColor$code[data$fillColor$studentNum == studentNum]
      } else {
        NULL
      }
      station_style <- if (!is.null(student_color) && student_color != "") {
        paste0("background-color:", student_color, ";")
      } else {
        ""
      }
      # Time label: merge all block times
      tb_times <- sapply(i:(i + span - 1), function(idx) {
        tb <- block_df$timeBlock[idx]
        if (!is.null(timeblock_times[[tb]])) timeblock_times[[tb]] else tb
      })
      time_label <- if (length(tb_times) == 1) tb_times else tagList(
        unlist(
          lapply(seq_along(tb_times), function(i) {
            if (i < length(tb_times)) list(tb_times[i], tags$br()) else tb_times[i]
          }),
          recursive = FALSE
        )
      )
      rows[[length(rows) + 1]] <- tags$tr(
        tags$td(time_label, rowspan = 1),
        tags$td(student_info, style = station_style, rowspan = 1)
      )
      i <- i + span
    }
  
    header <- tags$tr(
      tags$th("Time"),
      tags$th("Student Info")
    )
  
    tagList(
      tags$div(
        tags$h4(niceName),
        tags$p(tags$b("Group #:"), groupNum),
        if (!is.null(room1) && !is.na(room1) && room1 != "") tags$p(tags$b("Room:"), room1),
        if (!is.null(room2) && !is.na(room2) && room2 != "") tags$p(tags$b("Additional Room:"), room2),
        if (!faculty_by_student && !is.null(roomFaculty) && !is.na(roomFaculty) && roomFaculty != "") tags$p(tags$b("Faculty:"), roomFaculty),
        if (!is.null(duration) && !is.na(duration) && duration != "" && duration != 0) tags$p(tags$b("Duration:"), duration, "min"),
        tags$p(tags$b("Date:"), format(as.Date(group_date), "%A, %B %d, %Y")),
        tags$p(tags$b("Start time:"), format(strptime(format(as_hms(as.numeric(group_start) * 86400)), "%H:%M:%S"), "%I:%M %p")),
        tags$p(tags$b("End time:"), format(strptime(format(as_hms(as.numeric(group_end) * 86400)), "%H:%M:%S"), "%I:%M %p")),
        if (!is.null(notes) && !is.na(notes) && notes != "") tags$p(tags$b("Notes:"), notes)
      ),
      tags$table(
        id = "station_schedule_table",
        style = "border-collapse:collapse;width:100%;margin:auto;",
        tags$thead(header),
        tags$tbody(rows)
      ) %>%
        tagAppendChild(
          tags$style(HTML("
            #station_schedule_table tr th, #station_schedule_table tr td {
              border: 1px solid #333 !important;
              padding: 8px 12px !important;
            }
          "))
        )
    )
  })

  # Event Calendar Overview ---------------

  output$event_overview_calendar <- renderUI({
    req(data$groupInfo, data$timeBlockInfo)
    group_df <- data$groupInfo
    tb_info <- data$timeBlockInfo
  
    get_time_tod <- function(frac) {
      as.POSIXct(sprintf("%02d:%02d", floor(frac * 24), round((frac * 24 - floor(frac * 24)) * 60)), format="%H:%M", tz="UTC")
    }
  
    events <- lapply(seq_len(nrow(group_df)), function(i) {
      row <- group_df[i, ]
      tb_row <- which(tb_info$startTimeLabel == row$timeOfDay)
      if (length(tb_row) == 1) {
        arrival <- tb_info$arrivalTime[tb_row]
        leave <- tb_info$leaveTime[tb_row]
        list(
          groupNum = row$groupNum,
          date = as.character(row$date),
          start_tod = get_time_tod(arrival),
          end_tod = get_time_tod(leave)
        )
      } else {
        NULL
      }
    })
    events <- do.call(rbind, lapply(events, as.data.frame))
    if (is.null(events) || nrow(events) == 0) return(tags$div("No group events found."))
  
    all_dates <- sort(unique(events$date))
    min_tod <- min(events$start_tod)
    max_tod <- max(events$end_tod)
    time_seq <- seq(from = min_tod, to = max_tod, by = "15 min")
    time_labels <- format(time_seq, "%H:%M")
  
    date_columns <- list()
    for (d in all_dates) {
      day_events <- events[events$date == d, ]
      cols <- rep(NA_integer_, nrow(day_events))
      col_end <- c()
      for (i in order(day_events$start_tod)) {
        assigned <- FALSE
        for (j in seq_along(col_end)) {
          if (day_events$start_tod[i] >= col_end[j]) {
            cols[i] <- j
            col_end[j] <- day_events$end_tod[i]
            assigned <- TRUE
            break
          }
        }
        if (!assigned) {
          cols[i] <- length(col_end) + 1
          col_end <- c(col_end, day_events$end_tod[i])
        }
      }
      day_events$col <- cols
      date_columns[[d]] <- day_events
    }
    # Number of columns for each date
    ncols_by_date <- sapply(date_columns, function(df) max(df$col))
    total_cols <- sum(ncols_by_date)
    
    # Calculate the starting column for each date (for thick border)
    date_start_cols <- cumsum(c(2, head(ncols_by_date, -1)))
  
    # First header row: merged date headings
    header1 <- tags$tr(
      tags$th(rowspan = 2, "Time"),
      lapply(seq_along(all_dates), function(i) {
        # Format date as "Mon July 10"
        date_label <- format(as.Date(all_dates[i]), "%a %B %d")
        tags$th(date_label, colspan = ncols_by_date[i], style = "text-align:center;")
      })
    )
  
    # Track for each date/column which rows are covered by a rowspan
    covered <- lapply(all_dates, function(d) {
      ncol <- max(date_columns[[d]]$col)
      matrix(FALSE, nrow = length(time_seq), ncol = ncol)
    })
    names(covered) <- all_dates
  
    body_rows <- list()
    for (t in seq_along(time_seq)) {
      # Always show the first block's time label, and every hour
      hour_label <- if (t == 1 || format(time_seq[t], "%M") == "00") format(time_seq[t], "%H:%M") else ""
      row_cells <- list(tags$td(hour_label))
      col_counter <- 1
      for (d_idx in seq_along(all_dates)) {
        d <- all_dates[d_idx]
        day_events <- date_columns[[d]]
        ncol <- max(day_events$col)
        for (col_idx in seq_len(ncol)) {
          if (covered[[d]][t, col_idx]) next
          ev <- day_events[day_events$col == col_idx, ]
          found <- FALSE
          for (k in seq_len(nrow(ev))) {
            if (abs(as.numeric(difftime(time_seq[t], ev$start_tod[k], units = "mins"))) < 1) {
              span <- as.numeric(difftime(ev$end_tod[k], ev$start_tod[k], units = "mins")) / 15
              rows_to_cover <- t + seq_len(span) - 1
              rows_to_cover <- rows_to_cover[rows_to_cover <= nrow(covered[[d]])]
              covered[[d]][rows_to_cover, col_idx] <- TRUE
              # Add thick border class if this is the first col of a day (but not the first day)
              td_class <- if (col_idx == 1 && d_idx > 1) "thick-border-left" else NULL
              group_color <- if ("groupColor" %in% names(group_df)) {
                group_df$groupColor[group_df$groupNum == ev$groupNum[k]][1]
              } else {
                "#e0f7fa" # fallback color
              }
              row_cells[[length(row_cells) + 1]] <- tags$td(
                tags$div(
                  tags$b(input$event_nice_name),
                  tags$br(),
                  paste0("Group ", ev$groupNum[k]),
                  tags$br(),
                  paste0("(", format(ev$start_tod[k], "%H:%M"), " - ", format(ev$end_tod[k], "%H:%M"), ")")
                ),
                style = paste0("background:", group_color, ";font-weight:bold;text-align:center;"),
                rowspan = span,
                class = paste("event-block", td_class)
              )
              found <- TRUE
              break
            }
          }
          if (!found) {
            td_class <- if (col_idx == 1 && d_idx > 1) "thick-border-left" else NULL
            row_cells[[length(row_cells) + 1]] <- tags$td("", class = td_class)
          }
          col_counter <- col_counter + 1
        }
      }
      body_rows[[length(body_rows) + 1]] <- tags$tr(row_cells)
    }
    
    tags$table(
      id = "event_overview_calendar_table",
      style = "border-collapse:collapse;width:100%;margin:auto;",
      tags$thead(header1),
      tags$tbody(body_rows)
    ) %>%
      tagAppendChild(
        tags$style(HTML("
          #event_overview_calendar_table tr th, #event_overview_calendar_table tr td {
            border: 1px solid #333 !important;
            padding: 4px 8px !important;
          }
          #event_overview_calendar_table tr th {
            background: #f8f9fa;
            font-weight: bold;
            text-align: center;
          }
          #event_overview_calendar_table tr td {
            text-align: center;
          }
          .thick-border-left {
            border-left: 4px solid #333 !important;
          }
          /* Standardize row height for the Time column */
          #event_overview_calendar_table tr td:first-child {
            min-height: 28px;
            height: 28px;
          }
        "))
      )
  })


  ########################
  ## Download handlers
  ########################

  output$download <- downloadHandler(
    filename = function() {
      fname <- input$event_file_name
      if (is.null(fname) || fname == "") fname <- "Generated_Schedules"
      paste0(fname, ".xlsx")
    },
    content = function(file) {
      shiny::withProgress(
        message = "Saving schedules...",
        value = 0,
        {
          wb <- createWorkbook()
          sched_names <- names(data$schedules)
          n_scheds <- length(sched_names)
          n_steps <- 8 + n_scheds # 8 template sheets + n group schedules + 1 overview
  
          currentStep <- 1

          # --- Event Overview Calendar ---
          add_event_overview_calendar_sheet(wb, "Event Overview Calendar", data$groupInfo, data$timeBlockInfo, input$event_nice_name)
          shiny::setProgress(value = currentStep/n_steps, detail = "Event Overview Calendar")
          currentStep <- currentStep + 1

          # --- Group schedules ---
          write_group_schedules_to_workbook(
            wb = wb,
            schedules = data$schedules,
            fillColor = data$fillColor,
            n_steps = n_steps,
            progress_offset = currentStep - 1 # so currentStep is incremented for each group
          )
          currentStep <- currentStep + n_scheds
  
          # --- Add template sheets ---
          update_tmpl_starttime_names()
          update_tmpl_group_info()
          update_tmpl_station_info()
          update_faculty_assignments()
          update_tmpl_fillColor()
          tmpl <- template_data()
  
          addWorksheet(wb, "Event Info")
          event_info <- data.frame(
            Field = c("Event Name", "Faculty Contact"),
            Value = c(input$event_nice_name, input$event_faculty_contact),
            stringsAsFactors = FALSE
          )
          writeDataTable(wb, "Event Info", event_info, tableStyle = "TableStyleLight1")
          shiny::setProgress(value = currentStep/n_steps, detail = "Event Info sheet")
          currentStep <- currentStep + 1
  
          addWorksheet(wb, "studentInfo")
          writeDataTable(wb, "studentInfo", tmpl$studentInfo, tableStyle = "TableStyleLight1")
          shiny::setProgress(value = currentStep/n_steps, detail = "studentInfo sheet")
          currentStep <- currentStep + 1
  
          addWorksheet(wb, "groupInfo")
          writeDataTable(wb, "groupInfo", tmpl$groupInfo, tableStyle = "TableStyleLight1")
          shiny::setProgress(value = currentStep/n_steps, detail = "groupInfo sheet")
          currentStep <- currentStep + 1
  
          addWorksheet(wb, "fillColor")
          writeData(wb, "fillColor", tmpl$fillColor)
          for (i in seq_len(nrow(tmpl$fillColor))) {
            color <- tmpl$fillColor$code[i]
            if (!is.null(color) && color != "") {
              addStyle(
                wb, "fillColor",
                createStyle(fgFill = color),
                rows = i + 1,
                cols = 2,
                gridExpand = TRUE,
                stack = TRUE
              )
            }
          }
          shiny::setProgress(value = currentStep/n_steps, detail = "fillColor sheet")
          currentStep <- currentStep + 1
  
          addWorksheet(wb, "timeBlockInfo")
          writeDataTable(wb, "timeBlockInfo", tmpl$timeBlockInfo, tableStyle = "TableStyleLight1")
          time_style <- createStyle(numFmt = "hh:mm")
          time_cols <- which(grepl("Time$|_Start$|_End$", names(tmpl$timeBlockInfo)))
          addStyle(
            wb, "timeBlockInfo", time_style,
            rows = 2:(nrow(tmpl$timeBlockInfo) + 1),
            cols = time_cols,
            gridExpand = TRUE, stack = TRUE
          )
          shiny::setProgress(value = currentStep/n_steps, detail = "timeBlockInfo sheet")
          currentStep <- currentStep + 1
  
          addWorksheet(wb, "schedule")
          writeDataTable(wb, "schedule", tmpl$schedule, tableStyle = "TableStyleLight1")
          shiny::setProgress(value = currentStep/n_steps, detail = "schedule sheet")
          currentStep <- currentStep + 1
  
          addWorksheet(wb, "faculty")
          writeDataTable(wb, "faculty", tmpl$faculty, tableStyle = "TableStyleLight1")
          shiny::setProgress(value = currentStep/n_steps, detail = "faculty sheet")
          currentStep <- currentStep + 1
  
          saveWorkbook(wb, file, overwrite = TRUE)
          shiny::setProgress(value = 1, detail = "Done")
        }
      )
    }
  )

  write_group_schedules_to_workbook <- function(wb, schedules, fillColor, n_steps, progress_offset = 6) {
    # Use group order from groupInfo, fallback to names(schedules) if missing
    if (!is.null(data$groupInfo) && "groupNum" %in% names(data$groupInfo)) {
      sched_names <- paste0("Group_", as.character(data$groupInfo$groupNum))
      sched_names <- sched_names[sched_names %in% names(schedules)]
    } else {
      sched_names <- names(schedules)
    }
    for (idx in seq_along(sched_names)) {
      name <- sched_names[idx]
      sched <- schedules[[name]]
      ws_name <- substr(name, 1, 31)
  
      # Add worksheet for each group
      addWorksheet(wb, ws_name)
  
      # Prepare header row: Station + time labels
      timeblock_cols <- grep("^TimeBlock", names(sched$wide), value = TRUE)
      timeblock_times <- sched$timeblock_times
      header <- c("Station", sapply(timeblock_cols, function(tb) {
        if (!is.null(timeblock_times[[tb]])) timeblock_times[[tb]] else tb
      }))
  
      # Prepare data rows (do not merge yet)
      rows <- lapply(seq_len(nrow(sched$wide)), function(i) {
        row <- sched$wide[i, ]
        # Use stationColor if present
        station_color <- if ("stationColor" %in% names(row) && !is.na(row$stationColor) && row$stationColor != "") row$stationColor else NA
        station_info <- paste0(
          row$niceName,
          if (!is.null(row$room1) && !is.na(row$room1) && row$room1 != "") paste0("\nRoom: ", row$room1) else "",
          if (!is.null(row$room2) && !is.na(row$room2) && row$room2 != "") paste0("\nRoom: ", row$room2) else "",
          if (!is.null(row$faculty) && !is.na(row$faculty) && row$faculty != "") paste0("\nFaculty: ", row$faculty) else "Faculty: TBD",
          if (!is.null(row$notes) && !is.na(row$notes) && row$notes != "") paste0("\nNotes: ", row$notes) else ""
        )
        c(station_info, as.character(unlist(row[timeblock_cols])), station_color)
      })
      # Add stationColor as a hidden column for later styling
      df <- as.data.frame(do.call(rbind, rows), stringsAsFactors = FALSE)
      names(df) <- c(header, "stationColor__internal__")
  
      # Write date as a title row above the table
      writeData(wb, ws_name, paste0("Date: ", format(as.Date(sched$date), "%A, %B %d, %Y")), startRow = 1, startCol = 1)
      addStyle(wb, ws_name, createStyle(textDecoration = "bold", halign = "center", fontSize = 14), rows = 1, cols = 1, gridExpand = TRUE)
      mergeCells(wb, ws_name, cols = 1:(length(header)), rows = 1)
  
      # Write the table below the date (exclude the internal color column)
      writeData(wb, ws_name, df[, 1:length(header)], startRow = 3, startCol = 1, borders = "all", headerStyle = createStyle(textDecoration = "bold", border = "Bottom"))
  
      # Set column width for column A (Station column)
      setColWidths(wb, ws_name, cols = 1, widths = 40)
      setColWidths(wb, ws_name, cols = 2:length(header), widths = 26)
  
      # Wrap text for all data and header cells
      wrap_style <- createStyle(wrapText = TRUE)
      addStyle(wb, ws_name, wrap_style, rows = 3:(nrow(df) + 3), cols = 1:length(header), gridExpand = TRUE, stack = TRUE)
  
      # Merge adjacent cells with the same value for each row (timeblock columns only)
      for (i in seq_len(nrow(df))) {
        start_col <- 2 # first timeblock col
        end_col <- length(header)
        j <- start_col
        while (j <= end_col) {
          val <- df[i, j]
          run_start <- j
          while (j < end_col && df[i, j + 1] == val && val != "" && !is.na(val)) {
            j <- j + 1
          }
          if (j > run_start) {
            mergeCells(wb, ws_name, cols = run_start:j, rows = i + 3)
            # Only keep value in first cell, blank out others
            for (k in (run_start + 1):j) {
              writeData(wb, ws_name, "", startCol = k, startRow = i + 3)
            }
          }
          j <- j + 1
        }
      }
  
      # Add color formatting for station info cells (column 1)
      for (i in seq_len(nrow(df))) {
        scol <- df$stationColor__internal__[i]
        if (!is.na(scol) && scol != "") {
          addStyle(
            wb, ws_name,
            createStyle(fgFill = scol),
            rows = i + 3, cols = 1, gridExpand = TRUE, stack = TRUE
          )
        }
      }
  
      # Only color if there is a student name (". " in the value)
      for (col in seq_along(timeblock_cols)) {
        tb <- timeblock_cols[col]
        for (i in seq_len(nrow(sched$wide))) {
          val <- sched$wide[[tb]][i]
          studentNum <- NA
          color <- NULL
          if (!is.na(val) && val != "" && grepl("\\. ", val)) {
            matches <- regmatches(val, regexpr("^[0-9]+", val))
            if (length(matches) > 0 && matches != "") {
              studentNum <- as.integer(matches)
              if (!is.na(studentNum) && studentNum %in% fillColor$studentNum) {
                color <- fillColor$code[fillColor$studentNum == studentNum]
              }
            }
          }
          # Do not color if no student name (no ". " in val)
          if (!is.null(color)) {
            addStyle(
              wb, ws_name,
              createStyle(fgFill = color, halign = "center", textDecoration = NULL, fontColour = "#000000"),
              rows = i + 3, cols = col + 1, gridExpand = TRUE, stack = TRUE
            )
          }
        }
      }
  
      # Progress update
      shiny::setProgress(value = (progress_offset + idx)/n_steps, detail = paste("Saving", ws_name))
    }
  }

  add_event_overview_calendar_sheet <- function(wb, sheet_name = "Event overview calendar", group_df, tb_info, event_nice_name = "Event") {
    addWorksheet(wb, sheet_name)
    
    # Helper to convert fraction to POSIXct
    get_time_tod <- function(frac) {
      as.POSIXct(sprintf("%02d:%02d", floor(frac * 24), round((frac * 24 - floor(frac * 24)) * 60)), format="%H:%M", tz="UTC")
    }
    
    # Build events data.frame
    events <- lapply(seq_len(nrow(group_df)), function(i) {
      row <- group_df[i, ]
      tb_row <- which(tb_info$startTimeLabel == row$timeOfDay)
      if (length(tb_row) == 1) {
        arrival <- tb_info$arrivalTime[tb_row]
        leave <- tb_info$leaveTime[tb_row]
        list(
          groupNum = row$groupNum,
          date = as.character(row$date),
          start_tod = get_time_tod(arrival),
          end_tod = get_time_tod(leave),
          groupColor = if ("groupColor" %in% names(group_df)) row$groupColor else "#e0f7fa"
        )
      } else {
        NULL
      }
    })
    events <- do.call(rbind, lapply(events, as.data.frame))
    if (is.null(events) || nrow(events) == 0) {
      writeData(wb, sheet_name, "No group events found.")
      return(invisible(NULL))
    }
    
    all_dates <- sort(unique(events$date))
    min_tod <- as.POSIXct(format(min(events$start_tod), "%Y-%m-%d %H:00:00"), tz = "UTC")
    # Round max_tod up to the next 45-min mark (e.g., 16:45, 17:45, etc.)
    max_tod <- max(events$end_tod)
    if (!is.na(max_tod)) {
      mins <- as.numeric(format(max_tod, "%M"))
      if (mins < 45) {
        max_tod <- max_tod + (45 - mins) * 60
      } else if (mins > 45) {
        max_tod <- max_tod + (60 - mins + 45) * 60
      }
    }
    time_seq <- seq(from = min_tod, to = max_tod, by = "15 min")
    
    # Build date_columns structure (same as in Shiny)
    date_columns <- list()
    for (d in all_dates) {
      day_events <- events[events$date == d, ]
      cols <- rep(NA_integer_, nrow(day_events))
      col_end <- c()
      for (i in order(day_events$start_tod)) {
        assigned <- FALSE
        for (j in seq_along(col_end)) {
          if (day_events$start_tod[i] >= col_end[j]) {
            cols[i] <- j
            col_end[j] <- day_events$end_tod[i]
            assigned <- TRUE
            break
          }
        }
        if (!assigned) {
          cols[i] <- length(col_end) + 1
          col_end <- c(col_end, day_events$end_tod[i])
        }
      }
      day_events$col <- cols
      date_columns[[d]] <- day_events
    }
    ncols_by_date <- sapply(date_columns, function(df) max(df$col))
    total_cols <- sum(ncols_by_date)
    
    # --- 1. Merged title row ---
    writeData(wb, sheet_name, event_nice_name, startRow = 1, startCol = 1)
    mergeCells(wb, sheet_name, cols = 1:(total_cols + 1), rows = 1)
    addStyle(
      wb, sheet_name,
      createStyle(textDecoration = "bold", fontSize = 16, halign = "center"),
      rows = 1, cols = 1:(total_cols + 1), gridExpand = TRUE, stack = TRUE
    )
    
    # --- 2. Date range row ---
    date_range <- if (length(all_dates) == 1) {
      format(as.Date(all_dates[1]), "%B %d, %Y")
    } else {
      # If same year, show "March 10 - 14, 2025", else "March 30, 2025 - April 2, 2026"
      y1 <- format(as.Date(all_dates[1]), "%Y")
      y2 <- format(as.Date(all_dates[length(all_dates)]), "%Y")
      if (y1 == y2) {
        paste0(
          format(as.Date(all_dates[1]), "%B %d"),
          " - ",
          format(as.Date(all_dates[length(all_dates)]), "%d, %Y")
        )
      } else {
        paste0(
          format(as.Date(all_dates[1]), "%B %d, %Y"),
          " - ",
          format(as.Date(all_dates[length(all_dates)]), "%B %d, %Y")
        )
      }
    }
    writeData(wb, sheet_name, date_range, startRow = 2, startCol = 1)
    mergeCells(wb, sheet_name, cols = 1:(total_cols + 1), rows = 2)
    addStyle(
      wb, sheet_name,
      createStyle(textDecoration = "bold", fontSize = 12, halign = "center"),
      rows = 2, cols = 1:(total_cols + 1), gridExpand = TRUE, stack = TRUE
    )
    
    # --- 3. Faculty contact row ---
    faculty_contact <- paste0("Faculty Contact: ", input$event_faculty_contact)
    writeData(wb, sheet_name, faculty_contact, startRow = 3, startCol = 1)
    mergeCells(wb, sheet_name, cols = 1:(total_cols + 1), rows = 3)
    addStyle(
      wb, sheet_name,
      createStyle(fontSize = 11, halign = "center"),
      rows = 3, cols = 1:(total_cols + 1), gridExpand = TRUE, stack = TRUE
    )
    
    # --- 4. Header row (dates) ---
    header <- c("Time", rep("", total_cols))
    writeData(wb, sheet_name, t(header), startRow = 4, startCol = 1, colNames = FALSE)
    # Write date headings and merge across subcolumns
    col_counter <- 2
    for (i in seq_along(all_dates)) {
      date_label <- format(as.Date(all_dates[i]), "%a %b %d")
      ncols <- ncols_by_date[i]
      writeData(wb, sheet_name, date_label, startRow = 4, startCol = col_counter)
      if (ncols > 1) {
        mergeCells(wb, sheet_name, cols = col_counter:(col_counter + ncols - 1), rows = 4)
      }
      col_counter <- col_counter + ncols
    }
    
    date_start_cols <- cumsum(c(2, head(ncols_by_date, -1)))
    
    # --- 5. Body rows ---
    row_idx <- 5
    covered <- lapply(all_dates, function(d) {
      ncol <- max(date_columns[[d]]$col)
      matrix(FALSE, nrow = length(time_seq), ncol = ncol)
    })
    names(covered) <- all_dates
    
    for (t in seq_along(time_seq)) {
      hour_label <- if (format(time_seq[t], "%M") == "00") format(time_seq[t], "%H:%M") else ""
      row <- rep("", total_cols + 1)
      row[1] <- hour_label
      col_counter <- 2
      for (d_idx in seq_along(all_dates)) {
        d <- all_dates[d_idx]
        day_events <- date_columns[[d]]
        ncol <- max(day_events$col)
        for (col_idx in seq_len(ncol)) {
          if (covered[[d]][t, col_idx]) {
            col_counter <- col_counter + 1
            next
          }
          ev <- day_events[day_events$col == col_idx, ]
          found <- FALSE
          for (k in seq_len(nrow(ev))) {
            if (abs(as.numeric(difftime(time_seq[t], ev$start_tod[k], units = "mins"))) < 1) {
              row[col_counter] <- paste0(
                event_nice_name, "\n",
                "Group ", ev$groupNum[k], "\n",
                "(", format(ev$start_tod[k], "%H:%M"), " - ", format(ev$end_tod[k], "%H:%M"), ")"
              )
              span <- as.numeric(difftime(ev$end_tod[k], ev$start_tod[k], units = "mins")) / 15
              if (span > 1) {
                mergeCells(wb, sheet_name, cols = col_counter, rows = row_idx:(row_idx + span - 1))
              }
              addStyle(
                wb, sheet_name,
                createStyle(
                  fgFill = ev$groupColor[k],
                  halign = "center",
                  valign = "center",
                  textDecoration = "bold",
                  border = "TopBottomLeftRight",
                  wrapText = TRUE
                ),
                rows = row_idx:(row_idx + span - 1), cols = col_counter, gridExpand = TRUE, stack = TRUE
              )
              rows_to_cover <- t + seq_len(span) - 1
              rows_to_cover <- rows_to_cover[rows_to_cover <= nrow(covered[[d]])]
              covered[[d]][rows_to_cover, col_idx] <- TRUE
              found <- TRUE
              break
            }
          }
          col_counter <- col_counter + 1
        }
      }
      writeData(wb, sheet_name, t(row), startRow = row_idx, startCol = 1, colNames = FALSE)

      # Apply thick left border to all first subcolumns of each day (except the first day)
      if (length(date_start_cols) > 1) {
        # Add thick right border to all last subcolumns of each day (including the first)
        for (i in seq_along(ncols_by_date)) {
          col <- sum(ncols_by_date[seq_len(i)]) + 1
          addStyle(
            wb, sheet_name,
            createStyle(border = c("right"), borderStyle = c("thick")),
            rows = row_idx, cols = col, gridExpand = TRUE, stack = TRUE
          )
        }
      }
      row_idx <- row_idx + 1
    }
    
    # Style header
    addStyle(wb, sheet_name, createStyle(textDecoration = "bold", halign = "center", border = "Bottom"), rows = 4, cols = 1:(total_cols + 1), gridExpand = TRUE)
    setColWidths(wb, sheet_name, cols = 1:(total_cols + 1), widths = 18)
    setRowHeights(wb, sheet_name, rows = 5:(row_idx - 1), heights = 16)
  }

  # ---- Student Schedules Download Handler ----
  output$download_students <- downloadHandler(
    filename = function() {
      fname <- input$event_file_name
      if (is.null(fname) || fname == ""){
        fname <- "Student_Schedules"
      } else {
        fname <- paste0(fname, "_Student_Schedules")
      }
      paste0(fname, ".zip")
    },
    content = function(file) {
      shiny::withProgress(
        message = "Saving student schedules...",
        value = 0,
        {
          fname <- input$event_file_name
          if (is.null(fname) || fname == ""){
            fname <- "Student_Schedules"
          } else {
            fname <- paste0(fname, "_Student_Schedules")
          }
          tmpdir <- tempdir()
          group_files <- c()
          sched_names <- names(data$schedules)
          n_groups <- length(sched_names)
          total_students <- sum(sapply(sched_names, function(group_name) {
            sched <- data$schedules[[group_name]]
            groupNum <- as.character(gsub("^Group_", "", group_name))
            group_students <- data$studentInfo[data$studentInfo$groupNum == groupNum, ]
            nrow(group_students)
          }))
          student_counter <- 0
          group_counter <- 0
          for (group_name in sched_names) {
            sched <- data$schedules[[group_name]]
            groupNum <- as.character(gsub("^Group_", "", group_name))
            group_students <- data$studentInfo[data$studentInfo$groupNum == groupNum, ]
            if (nrow(group_students) == 0) next
            wb <- createWorkbook()
            for (i in seq_len(nrow(group_students))) {
              stu <- group_students[i, ]
              studentNum <- stu$studentNum
              lastName <- stu$lastName
              firstName <- stu$firstName
              ws_name <- paste0(studentNum, "_", substr(gsub("[^A-Za-z0-9]", "", lastName), 1, 12))
              ws_name <- substr(ws_name, 1, 31) # Excel sheet name limit
  
              # Get this student's schedule
              long_sched <- sched$long
              student_sched <- long_sched %>%
                filter(
                  studentNum == !!studentNum,
                  lastName == !!lastName,
                  firstName == !!firstName,
                  groupNum == !!groupNum
                ) %>%
                arrange(timeBlock)
  
              # --- SKIP if student_sched is empty ---
              if (nrow(student_sched) == 0) next
  
              # Prepare table: Time, Station Info
              timeblock_times <- sched$timeblock_times
              rows <- lapply(seq_len(nrow(student_sched)), function(j) {
                row <- student_sched[j, ]
                tb_time <- if (!is.null(timeblock_times[[row$timeBlock]])) timeblock_times[[row$timeBlock]] else row$timeBlock
                station_info <- paste0(
                  row$niceName,
                  if (!is.null(row$room1) && !is.na(row$room1) && row$room1 != "") paste0("\nRoom: ", row$room1) else "",
                  if (!is.null(row$room2) && !is.na(row$room2) && row$room2 != "") paste0("\nRoom: ", row$room2) else "",
                  if (!is.null(row$faculty) && !is.na(row$faculty) && row$faculty != "") paste0("\nFaculty: ", row$faculty) else "",
                  if (!is.null(row$notes) && !is.na(row$notes) && row$notes != "") paste0("\nNotes: ", row$notes) else ""
                )
                # Add stationColor for later styling
                scol <- if ("stationColor" %in% names(row) && !is.na(row$stationColor) && row$stationColor != "") row$stationColor else NA
                c(tb_time, station_info, scol)
              })
              df <- as.data.frame(do.call(rbind, rows), stringsAsFactors = FALSE)
              names(df) <- c("Time", "Station Info", "stationColor__internal__")
  
              addWorksheet(wb, ws_name)
  
              # Write student info as a title block
              group_date <- sched$date
              group_start <- sched$startTime
              group_end <- sched$endTime
              info_lines <- c(
                paste0(firstName, " ", lastName),
                paste0("Group #: ", groupNum),
                paste0("Student #: ", studentNum),
                paste0("Name: ", lastName, ", ", firstName),
                paste0("Date: ", format(as.Date(group_date), "%A, %B %d, %Y")),
                paste0("Start time: ", format(strptime(format(as_hms(as.numeric(group_start) * 86400)), "%H:%M:%S"), "%I:%M %p")),
                paste0("End time: ", format(strptime(format(as_hms(as.numeric(group_end) * 86400)), "%H:%M:%S"), "%I:%M %p"))
              )
              writeData(wb, ws_name, info_lines, startRow = 1, startCol = 1)
              addStyle(wb, ws_name, createStyle(textDecoration = "bold", fontSize = 12), rows = 1, cols = 1, gridExpand = TRUE)
              # Write the table below the info block (exclude color column)
              writeData(wb, ws_name, df[, 1:2], startRow = length(info_lines) + 2, startCol = 1, borders = "all", headerStyle = createStyle(textDecoration = "bold", border = "Bottom"))
              setColWidths(wb, ws_name, cols = 1, widths = 18)
              setColWidths(wb, ws_name, cols = 2, widths = 40)
              wrap_style <- createStyle(wrapText = TRUE)
              addStyle(wb, ws_name, wrap_style, rows = (length(info_lines) + 2):(nrow(df) + length(info_lines) + 2), cols = 1:2, gridExpand = TRUE, stack = TRUE)
  
              # Add color formatting for station info cells (column 2)
              for (r in seq_len(nrow(df))) {
                scol <- df$stationColor__internal__[r]
                if (!is.na(scol) && scol != "") {
                  addStyle(
                    wb, ws_name,
                    createStyle(fgFill = scol),
                    rows = r + length(info_lines) + 2, cols = 2, gridExpand = TRUE, stack = TRUE
                  )
                }
              }
              student_counter <- student_counter + 1
              shiny::setProgress(
                value = student_counter / total_students,
                detail = paste("Saving", firstName, lastName, "(Group", groupNum, ")")
              )
            }
            group_file <- file.path(tmpdir, paste0(fname, "_Group_", groupNum, ".xlsx"))
            saveWorkbook(wb, group_file, overwrite = TRUE)
            group_files <- c(group_files, group_file)
            group_counter <- group_counter + 1
          }
          # Zip all group files
          zip::zip(zipfile = file, files = group_files, mode = "cherry-pick")
          shiny::setProgress(value = 1, detail = "Done")
        }
      )
    }
  )

  # Download station schedules -----------------
    output$download_station_schedules <- downloadHandler(
      filename = function() {
        fname <- input$event_file_name
        if (is.null(fname) || fname == ""){
          fname <- "Station_Schedules"
        } else {
          fname <- paste0(fname, "_Station_Schedules")
        }
        paste0(fname, ".zip")
      },
      content = function(file) {
        shiny::withProgress(
          message = "Saving station schedules...",
          value = 0,
          {
            fname <- input$event_file_name
            if (is.null(fname) || fname == ""){
              fname <- "Station_Schedules"
            } else {
              fname <- paste0(fname, "_Station_Schedules")
            }
            tmpdir <- tempdir()
            group_files <- c()
            sched_names <- names(data$schedules)
            n_groups <- length(sched_names)
            for (group_idx in seq_along(sched_names)) {
              group_name <- sched_names[group_idx]
              sched <- data$schedules[[group_name]]
              groupNum <- as.character(gsub("^Group_", "", group_name))
              wb <- createWorkbook()
              # For each station in this group
              stations <- sched$wide$shortKey
              for (i in seq_along(stations)) {
                stationKey <- stations[i]
                station_wide <- sched$wide[sched$wide$shortKey == stationKey, ]
                station_long <- sched$long %>% filter(shortKey == stationKey, groupNum == groupNum) %>% arrange(timeBlock)
                # Compose a unique worksheet name: shortKey_niceName, max 31 chars, replace forbidden chars
                base_name <- paste0(station_wide$shortKey[1], "_", station_wide$niceName[1])
                ws_name <- gsub("[\\[\\]\\*\\?/\\\\]", "_", base_name) # Excel-forbidden chars
                ws_name <- substr(ws_name, 1, 31)
                # Ensure uniqueness within this workbook
                existing_names <- names(wb$worksheets)
                if (ws_name %in% existing_names) {
                  suffix <- 1
                  repeat {
                    try_name <- substr(paste0(ws_name, "_", suffix), 1, 31)
                    if (!(try_name %in% existing_names)) {
                      ws_name <- try_name
                      break
                    }
                    suffix <- suffix + 1
                  }
                }
                if (nrow(station_long) == 0) next
    
                # --- Merge logic: group by set of studentNums, merge consecutive blocks ---
                block_df <- station_long %>%
                  group_by(timeBlock) %>%
                  summarise(
                    studentNums = list(sort(unique(studentNum[!is.na(studentNum)]))),
                    studentLabels = list(studentLabel[!is.na(studentLabel) & studentLabel != ""]),
                    faculties = list(faculty[!is.na(faculty) & faculty != ""]),
                    .groups = "drop"
                  )
                block_df$merge_key <- sapply(block_df$studentNums, function(x) paste(sort(x), collapse = ","))
    
                rows <- list()
                n <- nrow(block_df)
                i_row <- 1
                while (i_row <= n) {
                  merge_key <- block_df$merge_key[i_row]
                  span <- 1
                  while (i_row + span <= n && block_df$merge_key[i_row + span] == merge_key) {
                    span <- span + 1
                  }
                  # Compose student info
                  studentLabels <- unique(unlist(block_df$studentLabels[i_row]))
                  faculties <- unique(unlist(block_df$faculties[i_row]))
                  if (length(studentLabels) > 0) {
                    student_info <- paste0("Student: ", paste(studentLabels, collapse = "; "))
                  } else {
                    student_info <- "Break"
                  }
                  if (sched$faculty_by_student && length(faculties) > 0) {
                    student_info <- paste0(student_info, "\nFaculty: ", paste(faculties, collapse = "; "))
                  }
                  # Color: use first studentNum if present
                  studentNum <- if (length(block_df$studentNums[[i_row]]) > 0) block_df$studentNums[[i_row]][1] else NA
                  student_color <- if (!is.na(studentNum) && studentNum %in% data$fillColor$studentNum) data$fillColor$code[data$fillColor$studentNum == studentNum] else NA
                  # Time label: merge all block times, newline separated
                  tb_times <- sapply(i_row:(i_row + span - 1), function(idx) {
                    tb <- block_df$timeBlock[idx]
                    if (!is.null(sched$timeblock_times[[tb]])) sched$timeblock_times[[tb]] else tb
                  })
                  time_label <- if (length(tb_times) == 1) tb_times else paste(tb_times, collapse = "\n")
                  rows[[length(rows) + 1]] <- c(time_label, student_info, student_color, span)
                  i_row <- i_row + span
                }
                df <- as.data.frame(do.call(rbind, rows), stringsAsFactors = FALSE)
                names(df) <- c("Time", "Student Info", "studentColor", "rowspan")
                df$rowspan <- as.integer(df$rowspan)
    
                addWorksheet(wb, ws_name)
                # Write station info as a header
                info_lines <- c(
                  paste0("Group #: ", groupNum),
                  paste0("Station: ", station_wide$niceName[1]),
                  if (!is.null(station_wide$room1[1]) && !is.na(station_wide$room1[1]) && station_wide$room1[1] != "") paste0("Room: ", station_wide$room1[1]) else NULL,
                  if (!is.null(station_wide$room2[1]) && !is.na(station_wide$room2[1]) && station_wide$room2[1] != "") paste0("Additional Room: ", station_wide$room2[1]) else NULL,
                  if (!is.null(station_wide$faculty[1]) && !is.na(station_wide$faculty[1]) && station_wide$faculty[1] != "") paste0("Faculty: ", station_wide$faculty[1]) else NULL,
                  if (!is.null(station_wide$notes[1]) && !is.na(station_wide$notes[1]) && station_wide$notes[1] != "") paste0("Notes: ", station_wide$notes[1]) else NULL,
                  if (!is.null(station_wide$timeInMin[1]) && !is.na(station_wide$timeInMin[1]) && station_wide$timeInMin[1] != "" && station_wide$timeInMin[1] != 0) paste0("Duration: ", station_wide$timeInMin[1], " min") else NULL,
                  if (!is.null(sched$date) && !is.na(sched$date)) paste0("Date: ", format(as.Date(sched$date), "%A, %B %d, %Y")) else NULL
                )
                writeData(wb, ws_name, info_lines, startRow = 1, startCol = 1)
                addStyle(wb, ws_name, createStyle(textDecoration = "bold", fontSize = 12), rows = 1, cols = 1, gridExpand = TRUE)
                # Write the table below the info block
                writeData(wb, ws_name, df[, 1:2], startRow = length(info_lines) + 2, startCol = 1, borders = "all", headerStyle = createStyle(textDecoration = "bold", border = "Bottom"))
                setColWidths(wb, ws_name, cols = 1, widths = 18)
                setColWidths(wb, ws_name, cols = 2, widths = 40)
                wrap_style <- createStyle(wrapText = TRUE)
                addStyle(wb, ws_name, wrap_style, rows = (length(info_lines) + 2):(nrow(df) + length(info_lines) + 2), cols = 1:2, gridExpand = TRUE, stack = TRUE)
    
                # Add color formatting for student info cells (column 2)
                for (r in seq_len(nrow(df))) {
                  scol <- df$studentColor[r]
                  if (!is.na(scol) && scol != "") {
                    addStyle(
                      wb, ws_name,
                      createStyle(fgFill = scol),
                      rows = r + length(info_lines) + 2, cols = 2, gridExpand = TRUE, stack = TRUE
                    )
                  }
                }
              }
              group_file <- file.path(tmpdir, paste0(fname, "_Group_", groupNum, ".xlsx"))
              saveWorkbook(wb, group_file, overwrite = TRUE)
              group_files <- c(group_files, group_file)
              shiny::setProgress(
                value = group_idx / n_groups,
                detail = paste("Saving Group", groupNum)
              )
            }
            zip::zip(zipfile = file, files = group_files, mode = "cherry-pick")
            shiny::setProgress(value = 1, detail = "Done")
          }
        )
      }
    )
  ########################
  ## Template Creator Logic
  ########################

  # --- Store template data as a reactive value ---
  template_data <- reactive({
    req(
      input$tmpl_num_groups,
      input$tmpl_max_students,
      input$tmpl_total_students,
      input$tmpl_num_timeblocks,
      input$tmpl_num_starttimes,
      input$tmpl_num_stations
    )

    num_groups <- input$tmpl_num_groups
    max_students <- input$tmpl_max_students
    total_students <- input$tmpl_total_students
    num_timeblocks <- input$tmpl_num_timeblocks
    num_starttimes <- input$tmpl_num_starttimes
    num_stations <- input$tmpl_num_stations

    # Get start time names from tmpl_inputs
    start_time_names <- sapply(seq_len(num_starttimes), function(i) {
      key <- paste0("tmpl_starttime_name_", i)
      nm <- tmpl_inputs$starttime_names[[key]]
      if (is.null(nm) || nm == "") paste0("Start", i) else nm
    })

    # Get time block times for each start time from tmpl_inputs
    timeblock_times <- lapply(seq_len(num_starttimes), function(st_idx) {
      sapply(seq_len(num_timeblocks), function(tb_idx) {
        key <- paste0("tmpl_timeblock_", st_idx, "_", tb_idx)
        val <- tmpl_inputs$timeblock_times[[key]]
        if (is.null(val)) "" else val
      })
    })
    names(timeblock_times) <- start_time_names

    # studentInfo
    studentInfo <- {
      df <- tmpl_students()
      if (is.null(df)) {
        data.frame(
          lastName = "",
          firstName = "",
          groupNum = rep(seq_len(num_groups), each = max_students, length.out = total_students),
          studentNum = seq_len(total_students),
          stringsAsFactors = FALSE
        )
      } else {
        df[, c("lastName", "firstName", "groupNum", "studentNum")]
      }
    }

    # groupInfo from tmpl_group_info
    groupInfo <- data.frame(
      groupNum = character(num_groups),
      date = as.Date(rep(NA, num_groups)),
      timeOfDay = character(num_groups),
      groupColor = character(num_groups),
      stringsAsFactors = FALSE
    )
    for (i in seq_len(num_groups)) {
      group <- tmpl_group_info$groups[[i]]
      groupInfo$groupNum[i] <- if (!is.null(group) && !is.null(group$groupNum)) group$groupNum else i
      groupInfo$date[i] <- if (!is.null(group) && !is.null(group$date)) group$date else ""
      groupInfo$timeOfDay[i] <- if (!is.null(group) && !is.null(group$timeOfDay)) {
        get_start_time_label(group$timeOfDay, start_time_names)
      } else {
        get_start_time_label(i, start_time_names)
      }
      groupInfo$groupColor[i] <- if (!is.null(group) && !is.null(group$groupColor)) {
        group$groupColor
      } else {
        "#e0f7fa"
      }
    }

    # fillColor: use actual color pickers, only up to max students
    fillColor <- data.frame(
      studentNum = seq_len(max_students),
      code = sapply(seq_len(max_students), function(i) {
        key <- paste0("tmpl_student_color_", i)
        val <- tmpl_fillColor$colors[[key]]
        if (is.null(val) || val == "") "#FFFFFF" else val
      }),
      stringsAsFactors = FALSE
    )

    # Build timeBlockInfo: one row per start time
    # Build timeBlockInfo: one row per start time, columns for each block's start/end
    timeBlockInfo <- data.frame(
      startTimeLabel = start_time_names,
      arrivalTime = NA_real_,
      leaveTime = NA_real_,
      stringsAsFactors = FALSE
    )

    for (i in seq_along(start_time_names)) {
      arrival_key <- paste0("tmpl_arrival_", i)
      end_key <- paste0("tmpl_end_", i)
      arrival_val <- tmpl_inputs$arrival_times[[arrival_key]]
      end_val <- tmpl_inputs$end_times[[end_key]]
      # Excel time as fraction of day (UTC, no timezone offset)
      timeBlockInfo$arrivalTime[i] <- if (!is.null(arrival_val) && !is.na(arrival_val)) {
        (hour(arrival_val) * 3600 + minute(arrival_val) * 60 + second(arrival_val)) / 86400
      } else {
        NA
      }
      timeBlockInfo$leaveTime[i] <- if (!is.null(end_val) && !is.na(end_val)) {
        (hour(end_val) * 3600 + minute(end_val) * 60 + second(end_val)) / 86400
      } else {
        NA
      }
      for (tb_idx in seq_len(num_timeblocks)) {
        start_key <- paste0("tmpl_timeblock_", i, "_", tb_idx, "_start")
        end_key <- paste0("tmpl_timeblock_", i, "_", tb_idx, "_end")
        start_val <- input[[start_key]]
        end_val <- input[[end_key]]
        timeBlockInfo[i, paste0("Block", tb_idx, "_Start")] <- if (!is.null(start_val) && !is.na(start_val)) {
          (hour(start_val) * 3600 + minute(start_val) * 60 + second(start_val)) / 86400
        } else {
          NA
        }
        timeBlockInfo[i, paste0("Block", tb_idx, "_End")] <- if (!is.null(end_val) && !is.na(end_val)) {
          (hour(end_val) * 3600 + minute(end_val) * 60 + second(end_val)) / 86400
        } else {
          NA
        }
      }
    }

    # schedule: use actual station info from UI
    schedule <- {
      n <- input$tmpl_num_stations
      stations <- tmpl_station_info$stations
      # Ensure all fields are present for each station
      required_fields <- c("shortKey", "niceName", "timeInMin", "room1", "room2", "notes", "stationColor")
      stations_filled <- lapply(seq_len(n), function(i) {
        s <- if (length(stations) >= i && !is.null(stations[[i]])) stations[[i]] else list()
        # Fill missing fields with defaults
        for (f in required_fields) {
          if (is.null(s[[f]])) {
            s[[f]] <- switch(f,
              shortKey = paste0("S", i),
              niceName = paste0("Station ", i),
              timeInMin = "",
              room1 = "",
              room2 = "",
              notes = "",
              stationColor = "#FFFFFF"
            )
          }
        }
        s
      })
      as.data.frame(do.call(rbind, lapply(stations_filled, as.data.frame)), stringsAsFactors = FALSE)
    }

    for (i in seq_len(num_timeblocks)) {
      schedule[[paste0("TimeBlock", i)]] <- ""
    }

    for (i in seq_len(input$tmpl_num_stations)) {
      for (j in seq_len(input$tmpl_num_timeblocks)) {
        inputId <- paste0("sched_", i, "_", j)
        val <- input[[inputId]]
        colname <- paste0("TimeBlock", j)
        # If multiple students are selected, store as comma-separated string
        if (!is.null(val)) {
          if (length(val) > 1) {
            schedule[i, colname] <- paste(val, collapse = ",")
          } else {
            schedule[i, colname] <- val
          }
        }
      }
    }

    # faculty: one row per station, one column per group, use station names/keys
    update_faculty_assignments()
    faculty <- NULL
    if (input$faculty_assign_mode == "room") {
      # By room: one row per station, columns for each group
      faculty <- data.frame(
        shortKey = schedule$shortKey,
        stringsAsFactors = FALSE
      )
      for (g in seq_len(num_groups)) {
        group_col <- paste0("group", g)
        faculty[[group_col]] <- sapply(seq_len(nrow(schedule)), function(i) {
          val <- NULL
          if (!is.null(faculty_assignments$by_room[[as.character(g)]])) {
            val <- faculty_assignments$by_room[[as.character(g)]][[as.character(i)]]
          }
          if (is.null(val)) "" else val
        })
      }
    } else {
      # By student: one row per studentNum in each group
      faculty <- data.frame(
        groupNum = character(),
        studentNum = integer(),
        faculty = character(),
        stringsAsFactors = FALSE
      )
      for (g in seq_len(num_groups)) {
        for (s in seq_len(max_students)) {
          val <- ""
          if (!is.null(faculty_assignments$by_student[[as.character(g)]])) {
            val <- faculty_assignments$by_student[[as.character(g)]][[as.character(s)]]
            if (is.null(val)) val <- ""
          }
          faculty <- rbind(
            faculty,
            data.frame(
              groupNum = as.character(g),
              studentNum = s,
              faculty = val,
              stringsAsFactors = FALSE
            )
          )
        }
      }
    }

    list(
      studentInfo = studentInfo,
      groupInfo = groupInfo,
      fillColor = fillColor,
      timeBlockInfo = timeBlockInfo,
      schedule = schedule,
      faculty = faculty
    )
  })

  observeEvent(input$load_info, {
    update_tmpl_starttime_names()
    update_tmpl_group_info()
    update_tmpl_station_info()
    update_faculty_assignments()
    update_tmpl_fillColor()
    tmpl <- template_data()
    data$studentInfo <- tmpl$studentInfo
    data$groupInfo <- tmpl$groupInfo
    data$fillColor <- tmpl$fillColor
    data$timeBlockInfo <- tmpl$timeBlockInfo
    data$schedule <- tmpl$schedule
    data$faculty <- tmpl$faculty
  })

  # this ensures that the UI elements for the template creator are always active, including when the tab hasn't been clicked yet
  outputOptions(output, "tmpl_starttime_names_ui", suspendWhenHidden = FALSE)
  outputOptions(output, "tmpl_timeblock_times_ui", suspendWhenHidden = FALSE)
  outputOptions(output, "tmpl_group_info_ui", suspendWhenHidden = FALSE)
  outputOptions(output, "tmpl_student_colors_ui", suspendWhenHidden = FALSE)
  outputOptions(output, "tmpl_student_overflow_warning", suspendWhenHidden = FALSE)
  outputOptions(output, "tmpl_student_warning_ui", suspendWhenHidden = FALSE)
  outputOptions(output, "tmpl_student_table", suspendWhenHidden = FALSE)
  outputOptions(output, "tmpl_station_info_ui", suspendWhenHidden = FALSE)
  outputOptions(output, "tmpl_schedule_ui", suspendWhenHidden = FALSE)
  outputOptions(output, "faculty_assignment_ui", suspendWhenHidden = FALSE)
}

shinyApp(ui, server)
