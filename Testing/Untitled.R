
library(shiny)
library(shinydashboard)
library(pool)
library(dplyr)
library(DBI)
library(RODBC)
library(DT)


########################################################################################################
# Set up direction 

dir <- "//ad-its.credit-agricole.fr/Amundi_Boston/Homedirs/buonomo/@Config/Desktop/Cat Bond Monitor/"
dir_sql <- "'\\\\ad-its.credit-agricole.fr\\Amundi_Boston\\Homedirs\\buonomo\\@Config\\Desktop\\Cat Bond Monitor\\bulk.csv'"

# Load the file of cleaning function
source("~/Desktop/AMUNDI DATABASE/ImprovedDatabase/Old Clean/BH_clean.R", local=TRUE)
source("~/Desktop/AMUNDI DATABASE/ImprovedDatabase/Old Clean/RBC_clean.R", local=TRUE)

########################################################################################################

# Based on the user's selection, choose corresponding cleaning function to apply
cleanFunc <- function(Broker, Format, new) {
  if (Broker == "BH") {
    if (Format == "CUSIP") {
      curr <- BH_long(new)
    } else if (Format == "Size") {
      curr <- BH_short(new)
    } else if (Format == "Name") {
      curr <- BH_long2(new)
    }
  } else if (Broker == "RBC") {
    curr <- RBC_clean(new)
  }
  return (curr)
}

########################################################################################################

# UI interface based on Dashboard
ui <- dashboardPage(
  dashboardHeader(title="CAT Bond Quote Monitor"),
  # Please ignore the warning raised because of the icon
  dashboardSidebar(
    collapsed = F, 
    div(htmlOutput("Welcome"), style = "padding: 20px"),
    sidebarMenu(
      menuItem("View Tables", tabName = "view_table", icon = icon("search")),
      menuItem("Insert Entries", tabName = "insert_value", icon = icon("edit")),
      menuItem("Update Tables", tabName = "update_table", icon = icon("exchange-alt")),
      menuItem("Aon Weekly ILS Pricing", tabName = "Aon_table", icon = icon("dollar")),
      menuItem("BeechHill Weekly ILS Pricing", tabName = "BH_table", icon = icon("dollar")),
      menuItem("GCS Weekly ILS Pricing", tabName = "GCS_table", icon = icon("dollar")),
      menuItem("RBC Weekly ILS Pricing", tabName = "RBC_table", icon = icon("dollar")),
      menuItem("SwissRe Weekly ILS Pricing", tabName = "SwissRe_table", icon = icon("dollar"))
    )
    
  ),
  
  # Contents in each tab, which tabName matches the name in sidebarMenu
  dashboardBody(
    tabItems(
      tabItem(tabName = "view_table",
              h2("Recent Market Quotes"),
              actionButton("refresh",label = "Refresh"),
              DT::dataTableOutput(outputId = "d1")),
      
      tabItem(tabName = "update_table", 
              h2("Manually Update Quotes"),
              actionButton("updated", "Update"),
              DT::dataTableOutput(outputId = "d2"),br(),
              verbatimTextOutput("test")),
      
      tabItem(tabName = "insert_value",
              fluidRow(
                sidebarPanel(
                  selectInput(
                    inputId = "Broker",
                    label="Choose a broker",
                    choices=list('', 'BH', 'RBC'),
                    selected=''),
                  conditionalPanel(
                    condition = "input.Broker == 'BH'",
                    selectInput("Format", "Select format of BH", choices = c('Name','CUSIP', 'Size')),
                    selected = NULL)
                ),
                
                mainPanel(
                  # Add Text
                  textAreaInput(inputId = "Long_Text", label = "TEXT:", rows = 15, resize = "both"), br(),
                  actionButton("update", "New Text"), br(),
                  DT::dataTableOutput("Text_Table")
                )
              )
      ),
      
      tabItem(tabName = "Aon_table",
              h2("This Weeks Aon ILS Pricing"),
              actionButton("refresh_Aon",label = "Refresh"),
              DT::dataTableOutput(outputId = "d3")
      ),
      
      tabItem(tabName = "BH_table",
              h2("This Weeks BeechHill ILS Pricing"),
              actionButton("refresh_BH",label = "Refresh"),
              DT::dataTableOutput(outputId = "d4")
      ),
      
      tabItem(tabName = "GCS_table",
              h2("This Weeks GCS ILS Pricing"),
              actionButton("refresh_GCS",label = "Refresh"),
              DT::dataTableOutput(outputId = "d5")
      ),
      tabItem(tabName = "RBC_table",
              h2("This Weeks RBC ILS Pricing"),
              actionButton("refresh_rbc",label = "Refresh"),
              DT::dataTableOutput(outputId = "d6")
      ),
      
      tabItem(tabName = "SwissRe_table",
              h2("This Weeks SwissRe ILS Pricing"),
              actionButton("refresh_swissre",label = "Refresh"),
              DT::dataTableOutput(outputId = "d7")
      )
    )
  )
)
# Server of R shiny
server <- function(input, output, session) {
  
  ### Connect to SQL
  db_conn <- odbcConnect("ILS")
  
  # Pre-load the data for "view_table"
  sql <- sqlQuery(channel=db_conn, 
                  paste("SELECT * FROM [dbo].[MarketQuotes] ORDER BY id DESC"), 
                  stringsAsFactors = FALSE)
  
  #Pre-load the data for "Aon_table"
  AonQuote <- sqlQuery(channel=db_conn, 
                       paste("SELECT * FROM dbo.Aon WHERE QuoteDate = (SELECT MAX (QuoteDate) FROM dbo.Aon)"), 
                       stringsAsFactors = FALSE)
  
  #Pre-load the data for "BH_table"
  BHQuote <- sqlQuery(channel=db_conn, 
                      paste("SELECT * FROM dbo.BH WHERE QuoteDate = (SELECT MAX (QuoteDate) FROM dbo.BH)"), 
                      stringsAsFactors = FALSE)
  
  #Pre-load the data for "GCS_table"
  GCSQuote <- sqlQuery(channel=db_conn, 
                       paste("SELECT * FROM dbo.GCS WHERE QuoteDate = (SELECT MAX (QuoteDate) FROM dbo.GCS)"), 
                       stringsAsFactors = FALSE)
  
  #Pre-load the data for "RBC_table"
  RBCQuote <- sqlQuery(channel=db_conn, 
                       paste("SELECT * FROM dbo.RBC WHERE QuoteDate = (SELECT MAX (QuoteDate) FROM dbo.RBC)"), 
                       stringsAsFactors = FALSE)
  
  #Pre-load the data for "SwissRe_table"
  swissQuote <- sqlQuery(channel=db_conn, 
                         paste("SELECT * FROM dbo.SwissRe WHERE IndicationDate = (SELECT MAX(IndicationDate) FROM dbo.SwissRe)"), 
                         stringsAsFactors = FALSE)
  
  # Retrive the column names from SQL, MarketQuotes table
  ColumnsOfTable <- sqlColumns(db_conn, "MarketQuotes")
  SqlName <- ColumnsOfTable$COLUMN_NAME
  
  
  ### Function for View Table
  RV <- reactiveValues(data = data.frame(sql))
  
  output$d1 <- DT::renderDT({RV$data})
  
  # Refresh and reload the data
  observeEvent(input$refresh, {
    RV$data <- sqlQuery(channel=db_conn, 
                        paste("SELECT * FROM (SELECT TOP 1000 * FROM [dbo].[MarketQuotes] ORDER BY id DESC) as Added"), 
                        stringsAsFactors = FALSE)	
  }, ignoreInit=T)
  
  observeEvent(input$refresh, {session$reload()})
  
  
  ### Function for Update Table
  
  # Json for dropdown list build in DT table
  js <- c(
    "function(settings){",
    "	$('#BrokerSelect').selectsize()",
    "	$('#ActionsSelect').selectsize()",
    "}"
  )
  
  # The data frame which will hold the updated result
  result <- reactiveValues( df = data.frame(QuoteDate = Sys.Date(), 
                                            Broker="hold", 
                                            CUSIP="CUSIP", 
                                            Bond="Bond", 
                                            Size=0, 
                                            Actions="hold",  
                                            Price=0))
  
  # The DT table with dropdown list shown in the page
  output$d2 <- DT::renderDataTable({
    
    df = data.frame(QuoteDate = Sys.Date(), 
                    Broker='<select Broker="" id="BrokerSelect">
                       		<option value="RBC">RBC</option>
                       		<option value="BeechHill">BeechHill</option>
                       		</select>', 
                    CUSIP="CUSIP", 
                    Bond="Bond", 
                    Size=0, 
                    Actions='<select Actions="" id="ActionsSelect">
                       		<option value="bid">bid</option>
                       		<option value="offer">offer</option>
                       		</select>',  
                    Price=0)
    
    
    DT::datatable(df, escape=FALSE, selection = 'none', editable = 'cell', rownames=FALSE,
                  options = list(
                    paging=FALSE, searching=FALSE, processing = FALSE,
                    initComplete = JS(js),
                    preDrawCallback = JS('function() { Shiny.unbindAll(this.api().table().node()); }'),
                    drawCallback = JS('function() { Shiny.bindAll(this.api().table().node()); } ')
                  )
    )
  }, server=TRUE)
  
  # The proxy used for reactive the data frame after user input
  proxy = dataTableProxy('d2', session=session)
  
  # Update the table triggered by user edit the cell
  observeEvent(input$d2_cell_edit, {
    info <- input$d2_cell_edit
    i = info$row
    j = info$col+1
    v = info$value
    
    result$df[i, j] <- v
    print(result$df)
  })
  
  # Once click the update, send the row to the SQL database
  observeEvent(input$updated, {
    # Insert row into SQL database
    db_conn <- odbcConnect("ILS")
    sqlQuery(db_conn, "Set Identity_Insert dbo.MarketQuotes On")
    last <- sqlQuery(channel=db_conn, 
                     "SELECT TOP 1 id FROM [dbo].[MarketQuotes] ORDER BY id DESC", 
                     stringsAsFactors = FALSE)
    id <- last[[1]]+1
    if (nchar(result$df$CUSIP) > 9) {
      result$df$CUSIP <- substr(result$df$CUSIP, 3, 11)
    }
    manual <- paste("INSERT INTO [dbo].[MarketQuotes]", 
                    "(", paste0(SqlName, collapse = ","), ")", 
                    "VALUES (",
                    paste("'", result$df$QuoteDate, "',", sep=""), 
                    paste("'", input$BrokerSelect, "',", sep=""),
                    paste("'", result$df$CUSIP, "',", sep=""),
                    paste("'", result$df$Bond, "',", sep=""),
                    paste(result$df$Size, ",", sep=""),
                    paste("'", input$ActionsSelect, "',", sep=""),
                    paste(result$df$Price, ",", sep=""),
                    paste(id, ")", sep=""),
                    sep="")
    sqlQuery(db_conn, manual)
  })
  
  
  ### Function for Insert Data
  # Generate Reactive Text Data
  Text_DF <- reactiveValues(data = data.frame(Text = character()))
  
  # Add New Text
  observeEvent(input$update, {
    # Combine New Text With Old Text
    Text_DF$data <- rbind(
      Text_DF$data, 
      data.frame(Text = str_replace_all(input$Long_Text, "\n", "<br>")) # Here the hack
    )
    # Everytime take the last block of data 
    new <- as.character(tail(Text_DF$data, n=1))
    
    curr <- data.frame(matrix(ncol = 7, nrow = 0))
    error <- ""
    
    # Avoid crash caused by enter the data not match the clean function
    tryCatch({  
      curr <- cleanFunc(input$Broker, input$Format, new)
    }, error = function(err){
      showNotification(paste("There is an error:", err), type = "error")
      error = 1
    })
    
    # Get the last id number in the SQL
    last <- sqlQuery(channel=db_conn, 
                     "SELECT TOP 1 id FROM [dbo].[MarketQuotes] ORDER BY id DESC", 
                     stringsAsFactors = FALSE)
    id <- data.frame(id=seq(from=last[[1]]+1, length.out=nrow(curr)))
    
    curr <- cbind(curr, id)
    curr$Price <- as.numeric(curr$Price)
    curr$id <- as.integer(curr$id)
    
    # Insert by BULK INSERT function in SQL, avoid the string format not match '' format in SQL
    write.csv(curr, paste(dir, "bulk.csv", sep=""), row.names=F)
    
    if (error != 1) {
      Query <- paste("BULK INSERT dbo.MarketQuotes FROM", dir_sql,
                     "WITH (FORMAT = 'CSV', FIRSTROW=2, FIELDTERMINATOR =',', ROWTERMINATOR ='\\n');")
      
      sqlQuery(db_conn, Query)
    }	
    
  })
  
  # Generate Table
  output$Text_Table = renderDataTable({
    Text_DF$data
  }, escape = FALSE)
  
  #Function for Aon Weekly Table
  RV_Aon <- reactiveValues(data = data.frame(AonQuote))
  
  output$d3 <- DT::renderDT({RV_Aon$data})
  
  observeEvent(input$refresh, {
    RV_Aon$data <- sqlQuery(channel=db_conn, 
                            paste("SELECT * FROM dbo.Aon WHERE QuoteDate = (SELECT MAX (QuoteDate) FROM dbo.Aon)"), 
                            stringsAsFactors = FALSE)	
  }, ignoreInit=T)
  
  observeEvent(input$refresh_Aon, {session$reload()})
  
  ### Function for BeechHill Weekly Table
  RV_BH <- reactiveValues(data = data.frame(BHQuote))
  
  output$d4 <- DT::renderDT({RV_BH$data})
  
  observeEvent(input$refresh, {
    RV_BH$data <- sqlQuery(channel=db_conn, 
                           paste("SELECT * FROM dbo.BH WHERE QuoteDate = (SELECT MAX (QuoteDate) FROM dbo.BH)"), 
                           stringsAsFactors = FALSE)	
  }, ignoreInit=T)
  
  observeEvent(input$refresh_BH, {session$reload()})
  
  ### Function for GCS Weekly Table
  RV_GCS <- reactiveValues(data = data.frame(GCSQuote))
  
  output$d5 <- DT::renderDT({RV_GCS$data})
  
  observeEvent(input$refresh, {
    RV_GCS$data <- sqlQuery(channel=db_conn, 
                            paste("SELECT * FROM dbo.GCS WHERE QuoteDate = (SELECT MAX (QuoteDate) FROM dbo.GCS)"), 
                            stringsAsFactors = FALSE)	
  }, ignoreInit=T)
  
  observeEvent(input$refresh_GCS, {session$reload()})
  
  ### Function for RBC Weekly Table
  RV_rbc <- reactiveValues(data = data.frame(RBCQuote))
  
  output$d6 <- DT::renderDT({RV_rbc$data})
  
  observeEvent(input$refresh, {
    RV_rbc$data <- sqlQuery(channel=db_conn, 
                            paste("SELECT * FROM dbo.RBC WHERE QuoteDate = (SELECT MAX (QuoteDate) FROM dbo.RBC)"), 
                            stringsAsFactors = FALSE)	
  }, ignoreInit=T)
  
  observeEvent(input$refresh_rbc, {session$reload()})
  
  
  ### Function for SwissRe Weekly Table
  RV_swissre <- reactiveValues(data = data.frame(swissQuote))
  
  output$d7 <- DT::renderDT({RV_swissre$data})
  
  observeEvent(input$refresh, {
    RV_swissre$data <- sqlQuery(channel=db_conn, 
                                paste("SELECT * FROM dbo.SwissRe WHERE IndicationDate = (SELECT MAX(IndicationDate) FROM dbo.SwissRe)"), 
                                stringsAsFactors = FALSE)	
  }, ignoreInit=T)	
  
  observeEvent(input$refresh_swissre, {session$reload()})
  
}


shinyApp(ui, server)

##################################################################################################################################

# odbcClose(db_conn)

##################################################################################################################################