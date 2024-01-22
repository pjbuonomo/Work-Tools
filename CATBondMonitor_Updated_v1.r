library(shiny)
library(shinydashboard)
library(DT)
library(dplyr)
library(DBI)
# Other required libraries...

# Load the file of cleaning function
source("~/Desktop/AMUNDI DATABASE/ImprovedDatabase/Old Clean/BH_clean.R", local=TRUE)
source("~/Desktop/AMUNDI DATABASE/ImprovedDatabase/Old Clean/RBC_clean.R", local=TRUE)

# UI interface based on Dashboard
ui <- dashboardPage(
  dashboardHeader(title = "CAT Bond Quote Monitor"),
  dashboardSidebar(
    collapsed = FALSE, 
    div(htmlOutput("Welcome"), style = "padding: 20px"),
    sidebarMenu(
      menuItem("View Tables", tabName = "view_table", icon = icon("search")),
      menuItem("Insert Entries", tabName = "insert_value", icon = icon("edit")),
      # Other menu items...
    )
  ),
  dashboardBody(
    tabItems(
      tabItem(tabName = "view_table",
              h2("Recent Market Quotes"),
              actionButton("refresh", label = "Refresh"),
              DT::dataTableOutput(outputId = "d1")),
      
      tabItem(tabName = "insert_value",
              fluidRow(
                sidebarPanel(
                  selectInput(inputId = "Broker", label = "Choose a broker",
                              choices = list('', 'BH', 'RBC'), selected = ''),
                  conditionalPanel(
                    condition = "input.Broker == 'BH'",
                    selectInput("Format", "Select format of BH", choices = c('Name','CUSIP', 'Size')),
                    selected = NULL
                  )
                ),
                mainPanel(
                  textAreaInput(inputId = "Long_Text", label = "TEXT:", rows = 15, resize = "both"), br(),
                  actionButton("update", "New Text"), br(),
                  DT::dataTableOutput("Text_Table")
                )
              )
      ),
      # Other tabItems...
    )
  )
)

# Server of R shiny
server <- function(input, output, session) {
  
  # Connect to SQL
  # db_conn <- odbcConnect("ILS")
  # ... [SQL database connection setup] ...
  
  # Initialize reactive value for Text_DF
  Text_DF <- reactiveVal(data.frame(Text = character()))
  
  # Live update for Text_Table based on Long_Text input
  observe({
    Text_DF(data.frame(Text = str_replace_all(input$Long_Text, "\n", "<br>")))
  })
  
  # Update Data and SQL on button click
  observeEvent(input$update, {
    new_text <- input$Long_Text
    processed_data <- tryCatch({
      cleanFunc(input$Broker, input$Format, new_text)
    }, error = function(e) {
      showNotification(paste("Error processing text:", e$message), type = "error")
      return(NULL)
    })
    
    # Check if data is valid and update
    if (!is.null(processed_data)) {
      # Append new data to the existing data frame
      # Update SQL database here if needed
      showNotification("Data processed and updated", type = "success")
    }
  })
  
  # Render Text Table
  output$Text_Table <- renderDataTable({
    Text_DF()
  }, escape = FALSE)
  
  # ... [Other server functionality] ...
}

shinyApp(ui, server)
