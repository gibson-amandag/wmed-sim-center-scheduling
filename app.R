library(shiny)
library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)

# Helper to load all sheets
load_data <- function(file) {
    list(
        studentInfo   = read.xlsx(file, sheet = "studentInfo", detectDates = TRUE),
        groupInfo     = read.xlsx(file, sheet = "groupInfo", detectDates = TRUE),
        fillColor     = read.xlsx(file, sheet = "fillColor", detectDates = TRUE),
        timeBlockInfo = read.xlsx(file, sheet = "timeBlockInfo", detectDates = TRUE),
        schedule      = read.xlsx(file, sheet = "schedule", detectDates = TRUE)
    )
}

# Generate group schedules
generate_group_schedules <- function(data) {
  schedules <- list()
  for (group in unique(data$studentInfo$groupNum)) {
    group_students <- data$studentInfo %>% filter(groupNum == group)
    group_meta     <- data$groupInfo %>% filter(groupNum == group)
    if (nrow(group_meta) == 0) next

    sched <- data$schedule
    time_blocks <- grep("^TimeBlock", names(sched), value = TRUE)
    tb_info <- data$timeBlockInfo

    # Wide version: replace studentNum with "studentNum. lastName, firstName"
    wide_sched <- sched
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
      sched,
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

# UI
ui <- fluidPage(
  titlePanel("OSCE Schedule Generator"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload Excel File", accept = ".xlsx"),
      downloadButton("download", "Download Schedules")
    ),
    mainPanel(
      tabsetPanel(
        tabPanel("Generated Schedules", uiOutput("scheduleTabs")),
        tabPanel("Student Schedule",
          selectInput("student_select", "Select Student", choices = NULL),
          uiOutput("student_schedule_table")
        ),
        tabPanel("Student Info", tableOutput("studentInfo")),
        tabPanel("Group Info", tableOutput("groupInfo")),
        tabPanel("Time Blocks", tableOutput("timeBlockInfo")),
        tabPanel("Schedule Template", uiOutput("schedule"))
      )
    )
  )
)

# Server
server <- function(input, output, session) {
  data <- reactiveValues()

  observeEvent(input$file, {
    req(input$file)
    tables <- load_data(input$file$datapath)
    data$studentInfo   <- tables$studentInfo
    data$groupInfo     <- tables$groupInfo
    data$fillColor     <- tables$fillColor
    data$timeBlockInfo <- tables$timeBlockInfo
    data$schedule      <- tables$schedule
  })

  output$studentInfo <- renderTable({
    req(data$studentInfo)
    df <- data$studentInfo
    df$groupNum <- as.integer(df$groupNum)
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
      style_str <- paste0("background-color:", prev_color, ";text-align:center;")
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
              color <- "#FFFFFF"
              textColor <- NULL
              if (!is.na(studentNum) && studentNum %in% data$fillColor$studentNum) {
                color <- data$fillColor$code[data$fillColor$studentNum == studentNum]
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
          style_str <- paste0("background-color:", prev_color, ";text-align:center;")
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
    groupNum <- as.integer(sel[1])
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
      filter(studentNum == studentNum, lastName == !!lastName, firstName == !!firstName, groupNum == !!groupNum)

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
        tags$h4("Student Schedule"),
        tags$p(tags$b("Group #:"), groupNum),
        tags$p(tags$b("Student #:"), studentNum),
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
}

shinyApp(ui, server)