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

    time_blocks   <- data$timeBlockInfo$timeBlock
    schedule_rows <- data$schedule

    long_sched <- schedule_rows %>%
      pivot_longer(
        cols = all_of(time_blocks),
        names_to = "timeBlock",
        values_to = "studentNum"
      ) %>%
      filter(!is.na(studentNum)) %>%
      left_join(group_students, by = "studentNum") %>%
      left_join(data$fillColor, by = "studentNum") %>%
      mutate(
        studentLabel = ifelse(
          !is.na(lastName),
          paste0(lastName, ", ", firstName, " (", code, ")"),
          NA
        )
      ) %>%
      select(
        timeBlock, shortKey, niceName, room1, room2, faculty, notes, studentLabel
      )

    schedules[[paste0("Group_", group)]] <- long_sched
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
        tabPanel("Generated Schedules", uiOutput("scheduleTabs"))
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
    req(data$schedules, data$groupInfo, data$timeBlockInfo)
    tabs <- lapply(names(data$schedules), function(name) {
      tabPanel(name, uiOutput(paste0("sched_", name)))
    })
    do.call(tabsetPanel, tabs)
  })

  observe({
    req(data$schedules, data$groupInfo, data$timeBlockInfo)
    lapply(names(data$schedules), function(name) {
      output[[paste0("sched_", name)]] <- renderUI({
        sched <- data$schedules[[name]]
        group_num <- as.integer(gsub("Group_", "", name))
        group_info <- data$groupInfo[data$groupInfo$groupNum == group_num, ]
        time_blocks <- data$timeBlockInfo

        # Identify time block columns
        timeblock_cols <- grep("^TimeBlock", names(data$schedule), value = TRUE)
        # Build table header with date and time
        header <- tags$tr(
          tags$th("Station"),
          lapply(seq_along(timeblock_cols), function(j) {
            tb_name <- timeblock_cols[j]
            tb_info <- time_blocks[j, ]
            date_str <- if (nrow(group_info) > 0 && !is.null(group_info$date)) as.character(group_info$date) else ""
            time_str <- if (nrow(tb_info) > 0 && !is.null(tb_info$time)) as.character(tb_info$time) else ""
            tags$th(
              tags$div(tb_name),
              tags$div(date_str),
              tags$div(time_str)
            )
          })
        )

        # Build table rows with merged cells for consecutive students
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
          prev_label <- NULL
          prev_color <- NULL
          prev_textColor <- NULL
          colspan <- 1
          cell_info <- list()
          for (j in seq_along(timeblock_cols)) {
            tb <- timeblock_cols[j]
            studentNum <- row[[tb]]
            # Find student info for label
            student_row <- data$studentInfo[data$studentInfo$studentNum == studentNum, ]
            if (nrow(student_row) > 0) {
              label <- paste0(
                student_row$studentNum, ". ",
                student_row$lastName, ", ",
                student_row$firstName
              )
            } else if (!is.na(studentNum) && studentNum != "") {
              label <- as.character(studentNum)
            } else {
              label <- "Break"
            }
            # Determine color
            if (!is.na(studentNum) && studentNum %in% data$fillColor$studentNum) {
              color <- data$fillColor$code[data$fillColor$studentNum == studentNum]
              textColor <- NULL
            } else if (is.na(studentNum) || studentNum == "") {
              color <- "#717171"
              textColor <- "white"
            } else {
              color <- "#FFFFFF"
              textColor <- NULL
            }
            # Merge logic
            if (j == 1) {
              prev_label <- label
              prev_color <- color
              prev_textColor <- if (exists("textColor")) textColor else NULL
              colspan <- 1
            } else if (identical(label, prev_label) && label != "Break") {
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
    })
  })

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