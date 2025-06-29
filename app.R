library(shiny)
library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)
library(zip)
library(colourpicker)

# Helper to load all sheets
load_data <- function(file) {
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

  faculty_long <- pivot_longer(
    data$faculty,
    cols = starts_with("group"),
    names_to = "groupNum",
    values_to = "faculty",
    names_prefix = "group"
  )
  # Ensure groupNum is character for join
  faculty_long$groupNum <- as.character(faculty_long$groupNum)

  for (group in unique(data$studentInfo$groupNum)) {
    group_students <- data$studentInfo %>% filter(groupNum == group)
    group_meta     <- data$groupInfo %>% filter(groupNum == group)
    if (nrow(group_meta) == 0) next

    sched <- data$schedule
    time_blocks <- grep("^TimeBlock", names(sched), value = TRUE)
    tb_info <- data$timeBlockInfo

    # --- Assign faculty for this group using faculty_long ---
    # Join sched with faculty_long by niceName and groupNum
    sched_with_faculty <- sched %>%
      left_join(
        faculty_long %>% filter(groupNum == as.character(group)),
        by = c("shortKey")
      ) %>%
      mutate(
        faculty = ifelse(!is.na(faculty), faculty, ifelse(!is.null(sched$faculty), sched$faculty, NA))
      )

    # Wide version: replace studentNum with "studentNum. lastName, firstName"
    wide_sched <- sched_with_faculty
    for (tb in time_blocks) {
      wide_sched[[tb]] <- sapply(wide_sched[[tb]], function(sn) {
        if (is.na(sn) || sn == "") return("")
        stu <- group_students[group_students$studentNum == sn, ]
        if (nrow(stu) > 0) {
          paste0(stu$studentNum, ". ", stu$lastName, ", ", stu$firstName)
        } else {
          as.character(sn)
        }
      })
    }

    # Long version: one row per station/time block (no date/time columns)
    long_sched <- tidyr::pivot_longer(
      sched_with_faculty,
      cols = all_of(time_blocks),
      names_to = "timeBlock",
      values_to = "studentNum"
    ) %>%
      left_join(group_students, by = "studentNum") %>%
      left_join(data$fillColor, by = "studentNum") %>%
      mutate(
        studentLabel = ifelse(
          !is.na(lastName),
          paste0(studentNum, ". ", lastName, ", ", firstName),
          as.character(studentNum)
        )
      )

    # Store group-level info and time block times as separate parameters
    group_date <- group_meta$date[1]
    group_startTime <- group_meta$startTime[1]
    group_endTime <- group_meta$endTime[1]
    group_timeOfDay <- if ("timeOfDay" %in% names(group_meta)) group_meta$timeOfDay[1] else NA

    # Pick correct time column for this group
    time_col <- if (!is.na(group_timeOfDay) && grepl("PM", group_timeOfDay, ignore.case = TRUE)) "pmTimes" else "amTimes"
    timeblock_times <- if (time_col %in% names(tb_info)) {
      setNames(as.character(tb_info[[time_col]]), tb_info$timeBlock)
    } else {
      setNames(rep(NA, length(tb_info$timeBlock)), tb_info$timeBlock)
    }

    schedules[[paste0("Group_", group)]] <- list(
      wide = wide_sched,
      long = long_sched,
      date = group_date,
      startTime = group_startTime,
      endTime = group_endTime,
      timeOfDay = group_timeOfDay,
      timeblock_times = timeblock_times
    )
  }
  return(schedules)
}

get_start_time_label <- function(index, start_time_names) {
  idx <- suppressWarnings(as.integer(index))
  if (!is.na(idx) && idx >= 1 && idx <= length(start_time_names)) {
    return(start_time_names[idx])
  } else {
    return(NA)
  }
}

# UI
ui <- fluidPage(
  titlePanel("OSCE Schedule Generator"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload Excel File", accept = ".xlsx"),
      downloadButton("download", "Download Schedules"),
      downloadButton("download_students", "Download Student Schedules")
    ),
    mainPanel(
      tabsetPanel(
        tabPanel("Template Creator",
          fluidRow(
            column(12, h3("Group information")),
            column(6, p("How many groups of students are you scheduling?")),
            column(6, numericInput("tmpl_num_groups", "# of groups", 2, min = 1))
          ),
          # --- NEW: Group info entry UI ---
          uiOutput("tmpl_group_info_ui"),
          fluidRow(
            column(6, p("What is the maximum number of students per group?")),
            column(6, numericInput("tmpl_max_students", "Max # of students/group", 6, min = 1))
          ),
          uiOutput("tmpl_student_colors_ui"),
          fluidRow(
            column(12, h3("Total number of students")),
            column(6, p("What is the total number of students across all groups?")),
            column(6, numericInput("tmpl_total_students", "Total # of students", 12, min = 1))
          ),
          fluidRow(
            column(12, h3("Time block information")),
            column(6, 
              p("How many time blocks are there in the schedule?"),
              em("This is the number of time slots for each station, and it can include breaks")
            ),
            column(6, numericInput("tmpl_num_timeblocks", "# of time blocks", 6, min = 1))
          ),
          fluidRow(
            column(12, h3("Start time information")),
            column(6, p("How many start times are there in the schedule?"),
              em("You might have different start times for different groups, e.g. AM/PM")
            ),
            column(6, numericInput("tmpl_num_starttimes", "# of startTimes", 2, min = 1))
          ),
          # --- NEW: Start time names and time block times ---
          uiOutput("tmpl_starttime_names_ui"),
          uiOutput("tmpl_timeblock_times_ui"),
          fluidRow(
            column(12, h3("Station information")),
            column(6, p("How many stations are there in the schedule?")),
            column(6, numericInput("tmpl_num_stations", "# of stations", 6, min = 1))
          ),
          downloadButton("download_template", "Download Blank Template")
        ),tabPanel("Generated Schedules", uiOutput("scheduleTabs")),
        tabPanel("Student Schedule",
          selectInput("student_select", "Select Student", choices = NULL),
          uiOutput("student_schedule_table")
        ),
        tabPanel("Student Info", tableOutput("studentInfo")),
        tabPanel("Group Info", tableOutput("groupInfo")),
        tabPanel("Time Blocks", tableOutput("timeBlockInfo")),
        tabPanel("Schedule Template", uiOutput("schedule")),
        # --- NEW TAB ---
        
      )
    )
  )
)

# Server
server <- function(input, output, session) {
  data <- reactiveValues()

  # --- NEW: Store template labels/times in reactiveValues ---
  tmpl_inputs <- reactiveValues(
    starttime_names = list(),
    timeblock_times = list()
  )

  # --- NEW: Store group info in reactiveValues ---
  tmpl_group_info <- reactiveValues(groups = list())

  # --- Observe and update start time labels ---
  observe({
    req(input$tmpl_num_starttimes)
    n <- input$tmpl_num_starttimes
    # Update stored labels if changed
    for (i in seq_len(n)) {
      key <- paste0("tmpl_starttime_name_", i)
      val <- input[[key]]
      if (!is.null(val)) tmpl_inputs$starttime_names[[key]] <- val
    }
    # Remove any extra labels if n decreased
    to_remove <- setdiff(names(tmpl_inputs$starttime_names), paste0("tmpl_starttime_name_", seq_len(n)))
    tmpl_inputs$starttime_names[to_remove] <- NULL
  })

  # --- Observe and update time block times ---
  observe({
    req(input$tmpl_num_timeblocks, input$tmpl_num_starttimes)
    n_tb <- input$tmpl_num_timeblocks
    n_st <- input$tmpl_num_starttimes
    isolate({
      for (st_idx in seq_len(n_st)) {
        for (tb_idx in seq_len(n_tb)) {
          key <- paste0("tmpl_timeblock_", st_idx, "_", tb_idx)
          val <- input[[key]]
          if (!is.null(val)) tmpl_inputs$timeblock_times[[key]] <- val
        }
      }
      # Remove any extra if n_tb or n_st decreased
      valid_keys <- unlist(lapply(seq_len(n_st), function(st_idx) {
        paste0("tmpl_timeblock_", st_idx, "_", seq_len(n_tb))
      }))
      to_remove <- setdiff(names(tmpl_inputs$timeblock_times), valid_keys)
      tmpl_inputs$timeblock_times[to_remove] <- NULL
    })
  })

  # --- Observe and update group info ---
  observe({
    req(input$tmpl_num_groups, input$tmpl_num_starttimes)
    n <- input$tmpl_num_groups
    isolate({
      # Save current values
      for (i in seq_len(n)) {
        prefix <- paste0("tmpl_group_", i, "_")
        groupNum <- input[[paste0(prefix, "groupNum")]]
        date <- input[[paste0(prefix, "date")]]
        startTime <- input[[paste0(prefix, "startTime")]]
        endTime <- input[[paste0(prefix, "endTime")]]
        timeOfDay <- input[[paste0(prefix, "timeOfDay")]]
        # print(paste("group", groupNum, "timeOfDay", timeOfDay, is.null(timeOfDay)))
        tmpl_group_info$groups[[i]] <- list(
          groupNum = if (!is.null(groupNum)) groupNum else paste0("Group ", i),
          date = if (!is.null(date)) date else NULL,
          startTime = if (!is.null(startTime)) startTime else "",
          endTime = if (!is.null(endTime)) endTime else "",
          timeOfDay = if (!is.null(timeOfDay)) timeOfDay else NA
        )
      }
      # Remove extra if n decreased
      if (length(tmpl_group_info$groups) > n) {
        tmpl_group_info$groups <- tmpl_group_info$groups[seq_len(n)]
      }
    })
  })

  # --- UI for start time names ---
  output$tmpl_starttime_names_ui <- renderUI({
    req(input$tmpl_num_starttimes)
    n <- input$tmpl_num_starttimes
    isolate({
      fluidRow(
        column(12, h4("Start time labels")),
        lapply(seq_len(n), function(i) {
          key <- paste0("tmpl_starttime_name_", i)
          val <- if (!is.null(tmpl_inputs$starttime_names[[key]])) tmpl_inputs$starttime_names[[key]]
                 else if (i == 1) "AM" else if (i == 2) "PM" else paste0("Start", i)
          column(6, textInput(key, paste0("Start time label ", i), value = val))
        })
      )
    })
  })

  # --- UI for time block times for each start time ---
  output$tmpl_timeblock_times_ui <- renderUI({
    req(input$tmpl_num_timeblocks, input$tmpl_num_starttimes)
    n_tb <- input$tmpl_num_timeblocks
    n_st <- input$tmpl_num_starttimes
    start_names <- sapply(seq_len(n_st), function(i) {
      key <- paste0("tmpl_starttime_name_", i)
      if (!is.null(tmpl_inputs$starttime_names[[key]])) tmpl_inputs$starttime_names[[key]] else if (i == 1) "AM" else if (i == 2) "PM" else paste0("Start", i)
    })
    tagList(
      lapply(seq_len(n_st), function(st_idx) {
        fluidRow(
          column(12, h5(paste0("Times for ", ifelse(!is.null(start_names[st_idx]) && start_names[st_idx] != "", start_names[st_idx], paste0("Start ", st_idx))))),
          lapply(seq_len(n_tb), function(tb_idx) {
            key <- paste0("tmpl_timeblock_", st_idx, "_", tb_idx)
            val <- if (!is.null(tmpl_inputs$timeblock_times[[key]])) tmpl_inputs$timeblock_times[[key]] else ""
            column(3, textInput(
              key,
              paste0("Time for Block ", tb_idx),
              value = val
            ))
          })
        )
      })
    )
  })

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
          # print(
          #   str(timeOfDay_val),
          #   str(get_start_time_label(timeOfDay_val, start_time_names))
          # )
          fluidRow(
            column(2, textInput(paste0(prefix, "groupNum"), paste0("Group ", i, " Name"), value = groupNum_val)),
            column(2, dateInput(paste0(prefix, "date"), "Date", value = if (!is.null(date_val)) date_val else NULL)),
            column(2, textInput(paste0(prefix, "startTime"), "Start Time (e.g. 08:00)", value = startTime_val)),
            column(2, textInput(paste0(prefix, "endTime"), "End Time (e.g. 12:00)", value = endTime_val)),
            column(2, selectInput(paste0(prefix, "timeOfDay"), "Time of Day", choices = time_of_day_choices, selected = timeOfDay_val))
          )
        })
      )
    })
  })

  # --- UI for student color pickers ---
  output$tmpl_student_colors_ui <- renderUI({
    req(input$tmpl_max_students)
    n <- input$tmpl_max_students
    # Default colors (repeat if needed)
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
          val <- input[[key]]
          default_val <- if (!is.null(val)) val else default_colors[(i - 1) %% length(default_colors) + 1]
          column(4, 
        tags$div(
          paste("Student", i),
          colourpicker::colourInput(key, NULL, value = default_val, showColour = "both")
        )
          )
        })
      )
    )
  })

  # Helper to load all sheets
  load_data <- function(file) {
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

    faculty_long <- pivot_longer(
      data$faculty,
      cols = starts_with("group"),
      names_to = "groupNum",
      values_to = "faculty",
      names_prefix = "group"
    )
    # Ensure groupNum is character for join
    faculty_long$groupNum <- as.character(faculty_long$groupNum)

    for (group in unique(data$studentInfo$groupNum)) {
      group_students <- data$studentInfo %>% filter(groupNum == group)
      group_meta     <- data$groupInfo %>% filter(groupNum == group)
      if (nrow(group_meta) == 0) next

      sched <- data$schedule
      time_blocks <- grep("^TimeBlock", names(sched), value = TRUE)
      tb_info <- data$timeBlockInfo

      # --- Assign faculty for this group using faculty_long ---
      # Join sched with faculty_long by niceName and groupNum
      sched_with_faculty <- sched %>%
        left_join(
          faculty_long %>% filter(groupNum == as.character(group)),
          by = c("shortKey")
        ) %>%
        mutate(
          faculty = ifelse(!is.na(faculty), faculty, ifelse(!is.null(sched$faculty), sched$faculty, NA))
        )

      # Wide version: replace studentNum with "studentNum. lastName, firstName"
      wide_sched <- sched_with_faculty
      for (tb in time_blocks) {
        wide_sched[[tb]] <- sapply(wide_sched[[tb]], function(sn) {
          if (is.na(sn) || sn == "") return("")
          stu <- group_students[group_students$studentNum == sn, ]
          if (nrow(stu) > 0) {
            paste0(stu$studentNum, ". ", stu$lastName, ", ", stu$firstName)
          } else {
            as.character(sn)
          }
        })
      }

      # Long version: one row per station/time block (no date/time columns)
      long_sched <- tidyr::pivot_longer(
        sched_with_faculty,
        cols = all_of(time_blocks),
        names_to = "timeBlock",
        values_to = "studentNum"
      ) %>%
        left_join(group_students, by = "studentNum") %>%
        left_join(data$fillColor, by = "studentNum") %>%
        mutate(
          studentLabel = ifelse(
            !is.na(lastName),
            paste0(studentNum, ". ", lastName, ", ", firstName),
            as.character(studentNum)
          )
        )

      # Store group-level info and time block times as separate parameters
      group_date <- group_meta$date[1]
      group_startTime <- group_meta$startTime[1]
      group_endTime <- group_meta$endTime[1]
      group_timeOfDay <- if ("timeOfDay" %in% names(group_meta)) group_meta$timeOfDay[1] else NA

      # Pick correct time column for this group
      time_col <- if (!is.na(group_timeOfDay) && grepl("PM", group_timeOfDay, ignore.case = TRUE)) "pmTimes" else "amTimes"
      timeblock_times <- if (time_col %in% names(tb_info)) {
        setNames(as.character(tb_info[[time_col]]), tb_info$timeBlock)
      } else {
        setNames(rep(NA, length(tb_info$timeBlock)), tb_info$timeBlock)
      }

      schedules[[paste0("Group_", group)]] <- list(
        wide = wide_sched,
        long = long_sched,
        date = group_date,
        startTime = group_startTime,
        endTime = group_endTime,
        timeOfDay = group_timeOfDay,
        timeblock_times = timeblock_times
      )
    }
    return(schedules)
  }

  observeEvent(input$file, {
    req(input$file)
    tables <- load_data(input$file$datapath)
    data$studentInfo   <- tables$studentInfo
    data$groupInfo     <- tables$groupInfo
    data$fillColor     <- tables$fillColor
    data$timeBlockInfo <- tables$timeBlockInfo
    data$schedule      <- tables$schedule
    data$faculty      <- tables$faculty
  })

  output$studentInfo <- renderTable({
    req(data$studentInfo)
    df <- data$studentInfo
    df$groupNum <- as.character(df$groupNum)
    df$studentNum <- as.integer(df$studentNum)
    df
  }, striped = TRUE, bordered = TRUE)

  output$groupInfo <- renderTable({
    req(data$groupInfo)
    df <- data$groupInfo
    df <- df %>%
        mutate(
            date = format(as.Date(date)),
            startTime = format(as_hms(startTime * 86400)),
            endTime = format(as_hms(endTime * 86400))
        )
    df
  }, striped = TRUE, bordered = TRUE)

  output$timeBlockInfo <- renderTable({
    req(data$timeBlockInfo)
    data$timeBlockInfo
  }, striped = TRUE, bordered = TRUE)

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

    # Build table rows with merged cells for consecutive studentNum
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
      prev_studentNum <- NULL
      colspan <- 1
      cell_info <- list()
      for (j in seq_along(timeblock_cols)) {
        tb <- timeblock_cols[j]
        studentNum <- as.integer(row[[tb]])
        # Determine color and label
        if (!is.na(studentNum) && studentNum %in% fill$studentNum) {
          color <- fill$code[fill$studentNum == studentNum]
          label <- as.character(studentNum)
          textColor <- NULL
        } else if (is.na(studentNum) || studentNum == "") {
          color <- "#717171"
          label <- "Break"
          textColor <- "white"
        } else {
          color <- "#FFFFFF"
          label <- as.character(studentNum)
          textColor <- NULL
        }
        # Merge logic
        if (j == 1) {
          prev_studentNum <- studentNum
          prev_label <- label
          prev_color <- color
          prev_textColor <- if (exists("textColor")) textColor else NULL
          colspan <- 1
        } else if (identical(studentNum, prev_studentNum) && label != "Break") {
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
          prev_studentNum <- studentNum
          prev_label <- label
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
      style = "border-collapse:collapse;width:100%;",
      tags$thead(header),
      tags$tbody(rows)
    ) %>%
      tagAppendChild(
      tags$style(HTML("
        table tr th, table tr td {
        border: 1px solid #333 !important;
        padding: 8px 12px !important;
        }
      "))
      )
  })

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
    req(data$schedules)
    tabs <- lapply(names(data$schedules), function(name) {
      tabPanel(name, uiOutput(paste0("sched_", name)))
    })
    do.call(tabsetPanel, tabs)
  })

  observe({
    req(data$schedules, data$fillColor)
    lapply(names(data$schedules), function(name) {
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
            # Extract studentNum for coloring
            studentNum <- NA
            cell_label <- val
            if (is.na(val) || val == "") {
              cell_label <- "Break"
              color <- "#717171"
              textColor <- "white"
            } else {
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
            }
            # Merge logic
            if (j == 1) {
              prev_studentNum <- studentNum
              prev_label <- cell_label
              prev_color <- color
              prev_textColor <- if (exists("textColor")) textColor else NULL
              colspan <- 1
            } else if (identical(studentNum, prev_studentNum) && cell_label != "Break") {
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
              prev_studentNum <- studentNum
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
            style = "border-collapse:collapse;width:100%;margin:auto;",
            tags$thead(header),
            tags$tbody(rows)
          ) %>%
            tagAppendChild(
              tags$style(HTML("
                table tr th, table tr td {
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
    if (length(sel) != 2) return(NULL)
    groupNum <- as.character(sel[1])
    studentNum <- as.integer(sel[2])
    selected_row <- data$studentInfo[data$studentInfo$groupNum == groupNum & data$studentInfo$studentNum == studentNum, ]
    if (nrow(selected_row) == 0) return(NULL)
    lastName <- selected_row$lastName[1]
    firstName <- selected_row$firstName[1]

    # Find the group schedule
    group_name <- paste0("Group_", groupNum)
    if (!group_name %in% names(data$schedules)) return(NULL)
    sched <- data$schedules[[group_name]]
    group_date <- sched$date
    group_start <- sched$startTime
    group_end <- sched$endTime
    timeblock_times <- sched$timeblock_times

    # Get all time blocks for this group
    long_sched <- sched$long
    student_sched <- long_sched %>%
      filter(studentNum == studentNum, lastName == !!lastName, firstName == !!firstName, groupNum == !!groupNum) %>%
      arrange(timeBlock)

    # If no schedule, return
    if (nrow(student_sched) == 0) return(tags$div("No schedule found for this student."))

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
        style = "border-collapse:collapse;width:100%;margin:auto;",
        tags$thead(header),
        tags$tbody(rows)
      ) %>%
        tagAppendChild(
          tags$style(HTML("
            table tr th, table tr td {
              border: 1px solid #333 !important;
              padding: 8px 12px !important;
            }
          "))
        )
    )
  })

  # ---- Download Handlers ----
  output$download <- downloadHandler(
    filename = function() {
      "Generated_Schedules.xlsx"
    },
    content = function(file) {
      wb <- createWorkbook()
      for (name in names(data$schedules)) {
        sched <- data$schedules[[name]]
        ws_name <- substr(name, 1, 31) # Excel sheet name limit

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

        # Optionally, add color formatting for student cells
        for (col in seq_along(timeblock_cols)) {
          tb <- timeblock_cols[col]
          for (i in seq_len(nrow(sched$wide))) {
            val <- sched$wide[[tb]][i]
            studentNum <- NA
            if (!is.na(val) && val != "") {
              matches <- regmatches(val, regexpr("^[0-9]+", val))
              if (length(matches) > 0 && matches != "") {
                studentNum <- as.integer(matches)
              }
            }
            color <- NULL
            if (!is.na(studentNum) && studentNum %in% data$fillColor$studentNum) {
              color <- data$fillColor$code[data$fillColor$studentNum == studentNum]
            } else if (is.na(val) || val == "") {
              color <- "#717171"
            }
            if (!is.null(color)) {
              addStyle(
                wb, ws_name,
                createStyle(fgFill = color, halign = "center", textDecoration = NULL, fontColour = ifelse(color == "#717171", "#FFFFFF", "#000000")),
                rows = i + 3, cols = col + 1, gridExpand = TRUE, stack = TRUE
              )
            }
          }
        }
      }
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )

  # ---- Student Schedules Download Handler ----
  output$download_students <- downloadHandler(
    filename = function() {
      "Student_Schedules.zip"
    },
    content = function(file) {
      tmpdir <- tempdir()
      group_files <- c()
      for (group_name in names(data$schedules)) {
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
            filter(studentNum == studentNum, lastName == !!lastName, firstName == !!firstName, groupNum == !!groupNum) %>%
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
              if (!is.null(row$faculty) && !is.na(row$faculty) && row$faculty != "") paste0("\nFaculty: ", row$faculty) else "Faculty: TBD",
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
        }
        group_file <- file.path(tmpdir, paste0("Group_", groupNum, "_Student_Schedules.xlsx"))
        saveWorkbook(wb, group_file, overwrite = TRUE)
        group_files <- c(group_files, group_file)
      }
      # Zip all group files
      zip::zip(zipfile = file, files = group_files, mode = "cherry-pick")
    }
  )

  # --- TEMPLATE CREATOR LOGIC ---
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
    studentInfo <- data.frame(
      studentNum = seq_len(total_students),
      groupNum = rep(seq_len(num_groups), each = max_students, length.out = total_students),
      lastName = "",
      firstName = "",
      stringsAsFactors = FALSE
    )

    # --- NEW: groupInfo from tmpl_group_info ---
    groupInfo <- data.frame(
      groupNum = character(num_groups),
      date = character(num_groups),
      startTime = character(num_groups),
      endTime = character(num_groups),
      timeOfDay = character(num_groups),
      stringsAsFactors = FALSE
    )
    for (i in seq_len(num_groups)) {
      group <- tmpl_group_info$groups[[i]]
      groupInfo$groupNum[i] <- if (!is.null(group) && !is.null(group$groupNum)) group$groupNum else paste0("Group ", i)
      groupInfo$date[i] <- if (!is.null(group) && !is.null(group$date)) group$date else ""
      groupInfo$startTime[i] <- if (!is.null(group) && !is.null(group$startTime)) group$startTime else ""
      groupInfo$endTime[i] <- if (!is.null(group) && !is.null(group$endTime)) group$endTime else ""
      groupInfo$timeOfDay[i] <- if (!is.null(group) && !is.null(group$timeOfDay)) group$timeOfDay else as.character(ifelse(i <= length(start_time_names), i, 1))    }

    # fillColor
    fillColor <- data.frame(
      studentNum = seq_len(total_students),
      code = "",
      stringsAsFactors = FALSE
    )

    # timeBlockInfo
    timeBlockInfo <- data.frame(
      timeBlock = paste0("TimeBlock", seq_len(num_timeblocks)),
      stringsAsFactors = FALSE
    )
    # Add a column for each start time name, with the entered times
    for (i in seq_along(start_time_names)) {
      timeBlockInfo[[start_time_names[i]]] <- timeblock_times[[i]]
    }

    # schedule
    schedule <- data.frame(
      niceName = paste0("Station ", seq_len(num_stations)),
      shortKey = paste0("S", seq_len(num_stations)),
      room1 = "",
      room2 = "",
      faculty = "",
      notes = "",
      stationColor = "",
      stringsAsFactors = FALSE
    )
    for (i in seq_len(num_timeblocks)) {
      schedule[[paste0("TimeBlock", i)]] <- ""
    }

    # faculty
    faculty <- data.frame(
      niceName = paste0("Station ", seq_len(num_stations)),
      shortKey = paste0("S", seq_len(num_stations)),
      stringsAsFactors = FALSE
    )
    for (g in seq_len(num_groups)) {
      faculty[[paste0("group", g)]] <- ""
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

  output$download_template <- downloadHandler(
    filename = function() {
      "Blank_OSCE_Template.xlsx"
    },
    content = function(file) {
      tmpl <- template_data()
      wb <- createWorkbook()
      addWorksheet(wb, "studentInfo")
      writeData(wb, "studentInfo", tmpl$studentInfo)
      addWorksheet(wb, "groupInfo")
      writeData(wb, "groupInfo", tmpl$groupInfo)
      addWorksheet(wb, "fillColor")
      writeData(wb, "fillColor", tmpl$fillColor)
      addWorksheet(wb, "timeBlockInfo")
      writeData(wb, "timeBlockInfo", tmpl$timeBlockInfo)
      addWorksheet(wb, "schedule")
      writeData(wb, "schedule", tmpl$schedule)
      addWorksheet(wb, "faculty")
      writeData(wb, "faculty", tmpl$faculty)
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)

# Future to-do:
# - fill color only up to max in group
# - update build code for time picker to not look for "amTimes/pmTimes"
# - don't add faculty to the schedule table
# - don't add nice name to the faculty table
# - add ability to enter information about stations within the template creator
# - add fill colors for stations
# - add option for faculty by room for faculty by student