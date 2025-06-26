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
      actionButton("generate", "Generate Schedules"),
      downloadButton("download", "Download Schedules")
    ),
    mainPanel(
      tabsetPanel(
        tabPanel("Student Info", tableOutput("studentInfo")),
        tabPanel("Group Info", tableOutput("groupInfo")),
        tabPanel("Time Blocks", tableOutput("timeBlockInfo")),
        tabPanel("Schedule Template", uiOutput("schedule")), # changed from tableOutput to uiOutput
        tabPanel("Generated Schedules", uiOutput("scheduleTabs")),
        tabPanel("Explore Group Schedules",
          selectInput("explore_group", "Select Group", choices = NULL),
          radioButtons("explore_format", "Format", choices = c("wide", "long"), inline = TRUE),
          tableOutput("explore_group_table")
        )
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
          paste0("Room 1: ", row$room1)
        },
        if (!is.null(row$room2) && !is.na(row$room2) && row$room2 != "") {
          list(tags$br(), paste0("Room 2: ", row$room2))
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
      cells <- list(tags$td(station_info))
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

  observeEvent(input$generate, {
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
              paste0("Room 1: ", row$room1)
            },
            if (!is.null(row$room2) && !is.na(row$room2) && row$room2 != "") {
              list(tags$br(), paste0("Room 2: ", row$room2))
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
          cells <- list(tags$td(station_info))
          for (tb in timeblock_cols) {
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
            style_str <- paste0("background-color:", color, ";text-align:center;")
            if (exists("textColor") && !is.null(textColor)) style_str <- paste0(style_str, "color:", textColor, ";")
            cells[[length(cells) + 1]] <- tags$td(
              cell_label,
              style = style_str
            )
            if (exists("textColor", inherits = FALSE)) rm(textColor, inherits = FALSE)
          }
          do.call(tags$tr, cells)
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

  # Show selected group schedule in selected format
  output$explore_group_table <- renderTable({
    req(data$schedules, input$explore_group, input$explore_format)
    scheds <- data$schedules[[input$explore_group]]
    if (input$explore_format == "wide") {
      scheds$wide
    } else {
      scheds$long
    }
  }, striped = TRUE, bordered = TRUE)

  output$download <- downloadHandler(
    filename = function() {
      "Generated_Schedules.xlsx"
    },
    content = function(file) {
      wb <- createWorkbook()
      for (name in names(data$schedules)) {
        addWorksheet(wb, name)
        writeData(wb, name, data$schedules[[name]])
      }
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)