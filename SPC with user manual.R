library(shiny)
library(ggplot2)
library(readxl)
library(dplyr)
library(qcc)

ui <- navbarPage("SPC Dashboard",
                 tabPanel("Dashboard",
                          fluidPage(
                            titlePanel("SPC Chart Generator"),
                            
                            sidebarLayout(
                              sidebarPanel(
                                fileInput("file", "Upload Excel File:", accept = c(".xlsx", ".xls")),
                                uiOutput("sheet_selector"),
                                uiOutput("var_selector"),
                                uiOutput("group_selector"),
                                radioButtons("chart_type", "Select SPC Chart Type:",
                                             choices = c("Levey-Jennings", "IMR", "X-bar R", "np Chart", "p Chart", "c Chart", "u Chart")),
                                numericInput("group_size", "Group/Subgroup Size (only for attribute charts):", value = 5, min = 1),
                                downloadButton("download_chart", "Download Chart")
                              ),
                              
                              mainPanel(
                                plotOutput("spcPlot", height = "650px"),
                                verbatimTextOutput("summaryText"),
                                tableOutput("previewData"),
                                conditionalPanel(
                                  condition = "input.chart_type == 'Levey-Jennings' || input.chart_type == 'IMR'",
                                  tags$h4("Example Dummy Data for Levey-Jennings / IMR"),
                                  tableOutput("dummy_lj")
                                ),
                                conditionalPanel(
                                  condition = "input.chart_type == 'X-bar R'",
                                  tags$h4("Example Dummy Data for X-bar R Chart"),
                                  tableOutput("dummy_xbar")
                                ),
                                conditionalPanel(
                                  condition = "input.chart_type == 'np Chart' || input.chart_type == 'p Chart' || input.chart_type == 'c Chart' || input.chart_type == 'u Chart'",
                                  tags$h4("Example Dummy Data for Attribute Charts (np/p/c/u)"),
                                  tableOutput("dummy_attr")
                                )
                              )
                            )
                          )
                 ),
                 
                 tabPanel("User Panel",
                          fluidPage(
                            titlePanel("User Manual"),
                            tags$div(style = "padding: 20px; font-size: 16px;",
                                     HTML("
          <h3>SPC Dashboard User Manual</h3>
          <p><strong>Purpose:</strong> This app allows users to generate SPC (Statistical Process Control) charts from uploaded Excel data.</p>

          <h4>Steps to Use:</h4>
          <ol>
            <li><strong>Upload Excel File</strong>: Upload <code>.xlsx</code> or <code>.xls</code> file with relevant numeric data.</li>
            <li><strong>Select Sheet</strong>: Choose the sheet containing the data.</li>
            <li><strong>Select Main Variable</strong>: Choose the numeric variable for charting.</li>
            <li><strong>Choose Chart Type</strong>: Select from Levey-Jennings, IMR, X-bar R, np, p, c, or u chart.</li>
            <li><strong>Enter Group/Subgroup Size</strong> (for attribute charts).</li>
            <li><strong>Download Chart</strong>: Click the button to save the chart as PNG.</li>
          </ol>

          <h4>Chart Descriptions:</h4>
          <ul>
            <li><strong>Levey-Jennings</strong>: For lab/control measures with ±SD lines.</li>
            <li><strong>IMR</strong>: For individual observations over time.</li>
            <li><strong>X-bar R</strong>: For subgrouped continuous data (n > 1).</li>
            <li><strong>np/p Chart</strong>: For defectives in constant/varying sized samples.</li>
            <li><strong>c/u Chart</strong>: For defects per unit, with constant/varying areas.</li>
          </ul>

          <p><em>Note:</em> Example dummy data is shown under each chart type on the Dashboard tab.</p>
        ")
                            )
                          )
                 )
)

server <- function(input, output, session) {
  
  output$sheet_selector <- renderUI({
    req(input$file)
    sheets <- excel_sheets(input$file$datapath)
    selectInput("sheet", "Select Sheet:", choices = sheets)
  })
  
  dataInput <- reactive({
    req(input$file, input$sheet)
    read_excel(input$file$datapath, sheet = input$sheet)
  })
  
  output$var_selector <- renderUI({
    req(dataInput())
    selectInput("var", "Select Main Variable:", choices = names(dataInput()))
  })
  
  output$group_selector <- renderUI({
    req(dataInput())
    if (input$chart_type %in% c("X-bar R", "np Chart", "p Chart", "c Chart", "u Chart")) {
      selectInput("group", "Select Group Variable:", choices = names(dataInput()))
    }
  })
  
  output$previewData <- renderTable({
    dataInput()
  })
  
  chartSummary <- reactive({
    switch(input$chart_type,
           "Levey-Jennings" = "Use Levey-Jennings chart for lab or quality control measurements around a central mean.",
           "IMR" = "Use Individual-Moving Range (IMR) chart when you have single observations over time (n=1 per subgroup).",
           "X-bar R" = "Use X-bar R chart for subgrouped continuous data (n>1). Tracks both mean and variability.",
           "np Chart" = "Use np chart when tracking number of defectives in constant-sized samples.",
           "p Chart" = "Use p chart when tracking proportion of defectives in varying-sized samples.",
           "c Chart" = "Use c chart for count of defects per unit with constant inspection area.",
           "u Chart" = "Use u chart for defects per unit when inspection area varies."
    )
  })
  
  output$summaryText <- renderText({
    chartSummary()
  })
  
  chartPlot <- reactive({
    req(input$var)
    data <- dataInput()
    x <- unlist(data[[input$var]])
    
    qcc.options(
      title.font = 2, title.cex = 1.6,
      labels.cex = 1.4, cex = 1.4,
      cex.axis = 1.4, cex.lab = 1.4
    )
    
    if (input$chart_type == "Levey-Jennings") {
      mean_val <- mean(x, na.rm = TRUE)
      sd_val <- sd(x, na.rm = TRUE)
      df <- data.frame(Value = x, Index = seq_along(x))
      ggplot(df, aes(x = Index, y = Value)) +
        geom_line(color = "#0072B2", size = 1.2) +
        geom_point(size = 3, color = "#D55E00") +
        geom_hline(yintercept = mean_val, color = "blue", linetype = "dashed", size = 1) +
        geom_hline(yintercept = mean_val + sd_val, color = "red", linetype = "dotted", size = 1) +
        geom_hline(yintercept = mean_val - sd_val, color = "red", linetype = "dotted", size = 1) +
        geom_hline(yintercept = mean_val + 2 * sd_val, color = "orange", linetype = "dotted", size = 1) +
        geom_hline(yintercept = mean_val - 2 * sd_val, color = "orange", linetype = "dotted", size = 1) +
        geom_hline(yintercept = mean_val + 3 * sd_val, color = "darkred", linetype = "dotted", size = 1) +
        geom_hline(yintercept = mean_val - 3 * sd_val, color = "darkred", linetype = "dotted", size = 1) +
        labs(title = "Levey-Jennings Chart", y = input$var, x = "Observation Number") +
        theme_minimal(base_size = 16) +
        theme(plot.title = element_text(color = "#333333", face = "bold", size = 20))
    } else if (input$chart_type == "IMR") {
      qcc(x, type = "xbar.one", title = "IMR Chart")
    } else if (input$chart_type == "X-bar R") {
      req(input$group)
      values <- matrix(unlist(data[input$var]), ncol = input$group_size, byrow = TRUE)
      qcc(values, type = "xbar", title = "X-bar R Chart")
    } else if (input$chart_type == "np Chart") {
      req(input$group)
      counts <- table(data[[input$group]])
      qcc(as.numeric(counts), type = "np", size = input$group_size, title = "np Chart")
    } else if (input$chart_type == "p Chart") {
      req(input$group)
      df <- data %>% group_by(.data[[input$group]]) %>% summarise(Defectives = sum(.data[[input$var]]), n = input$group_size)
      qcc(df$Defectives, type = "p", size = df$n, title = "p Chart")
    } else if (input$chart_type == "c Chart") {
      qcc(x, type = "c", title = "c Chart")
    } else if (input$chart_type == "u Chart") {
      req(input$group)
      df <- data %>% group_by(.data[[input$group]]) %>% summarise(Defects = sum(.data[[input$var]]), Units = input$group_size)
      qcc(df$Defects, type = "u", size = df$Units, title = "u Chart")
    }
  })
  
  output$spcPlot <- renderPlot({
    chartPlot()
  })
  
  output$download_chart <- downloadHandler(
    filename = function() {
      paste("SPC_Chart_", input$chart_type, "_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      png(file, width = 800, height = 600)
      print(chartPlot())
      dev.off()
    }
  )
  
  output$dummy_lj <- renderTable({
    data.frame(Measurement = c(101, 98, 99))
  })
  
  output$dummy_xbar <- renderTable({
    data.frame(Batch = c("A", "A", "B", "B"), Measurement = c(98, 100, 97, 102))
  })
  
  output$dummy_attr <- renderTable({
    data.frame(Day = 1:3, Defectives = c(2, 1, 3))
  })
}

shinyApp(ui = ui, server = server)
