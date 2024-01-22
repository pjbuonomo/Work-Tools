library(shiny)
library(shinydashboard)
library(RODBC)
library(DT)
library(dplyr)
library(DBI)

# Set up directories
dir <- "//ad-its.credit-agricole.fr/Amundi_Boston/Homedirs/buonomo/@Config/Desktop/Cat Bond Monitor/"
dir_sql <- "'\\\\ad-its.credit-agricole.fr\\Amundi_Boston\\Homedirs\\buonomo\\@Config\\Desktop\\Cat Bond Monitor\\bulk.csv'"

# Load the file of cleaning function
source("S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/BH_clean.R", local = TRUE)
source("S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/RBC_clean.R", local = TRUE)

# Server of R shiny
server <- function(input, output, session) {
    # Connect to SQL
    db_conn <- odbcConnect("ILS")

    # Pre-load the data for "view_table"
    sql <- sqlQuery(channel = db_conn, paste("SELECT * FROM [dbo].[MarketQuotes] ORDER BY id DESC"), stringsAsFactors = FALSE)
    
    # Pre-load the data for other tables
    AonQuote <- sqlQuery(channel = db_conn, paste("SELECT * FROM dbo.Aon WHERE QuoteDate = (SELECT MAX(QuoteDate) FROM dbo.Aon)"), stringsAsFactors = FALSE)
    BHQuote <- sqlQuery(channel = db_conn, paste("SELECT * FROM dbo.BH WHERE QuoteDate = (SELECT MAX(QuoteDate) FROM dbo.BH)"), stringsAsFactors = FALSE)
    GCSQuote <- sqlQuery(channel = db_conn, paste("SELECT * FROM dbo.GCS WHERE QuoteDate = (SELECT MAX(QuoteDate) FROM dbo.GCS)"), stringsAsFactors = FALSE)
    RBCQuote <- sqlQuery(channel = db_conn, paste("SELECT * FROM dbo.RBC WHERE QuoteDate = (SELECT MAX(QuoteDate) FROM dbo.RBC)"), stringsAsFactors = FALSE)
    swissQuote <- sqlQuery(channel = db_conn, paste("SELECT * FROM dbo.SwissRe WHERE IndicationDate = (SELECT MAX(IndicationDate) FROM dbo.SwissRe)"), stringsAsFactors = FALSE)

    # Retrieve the column names from SQL, MarketQuotes table
    ColumnsOfTable <- sqlColumns(db_conn, "MarketQuotes")
    SqlName <- ColumnsOfTable$COLUMN_NAME

    # Function for View Table
    RV <- reactiveValues(data = data.frame(sql))
    output$d1 <- DT::renderDT({
        RV$data
    })

    # Refresh and reload the data
    observeEvent(input$refresh, {
        RV$data <- sqlQuery(channel = db_conn, paste("SELECT * FROM (SELECT TOP 1000 * FROM [dbo].[MarketQuotes] ORDER BY id DESC) as Added"), stringsAsFactors = FALSE)
    }, ignoreInit = TRUE)
    
    observeEvent(input$refresh, {
        session$reload()
    })

    # (Continued in the next part...)
}

# (UI definition will be in the second part)

# Run the application (this will be at the end of the second part)
# shinyApp(ui = ui, server = server)
    # ... (Continuation from the previous part)

    ### Function for Update Table
    # JSON for dropdown list build in DT table
    js <- c(
      "function(settings){",
      " $('#BrokerSelect').selectsize()",
      " $('#ActionsSelect').selectsize()",
      "}"
    )

    # The data frame which will hold the updated result
    result <- reactiveValues(df = data.frame(QuoteDate = Sys.Date(), Broker = "hold", CUSIP = "CUSIP", Bond = "Bond", Size = 0, Actions = "hold", Price = 0))

    # The DT table with dropdown list shown in the page
    output$d2 <- DT::renderDataTable({
      df <- data.frame(QuoteDate = Sys.Date(), Broker = '<select Broker="" id="BrokerSelect"> <option value="RBC">RBC</option> <option value="BeechHill">BeechHill</option> </select>', CUSIP = "CUSIP", Bond = "Bond", Size = 0, Actions = '<select Actions="" id="ActionsSelect"> <option value="bid">bid</option> <option value="offer">offer</option> </select>', Price = 0)
      DT::datatable(df, escape = FALSE, selection = 'none', editable = 'cell', rownames = FALSE, options = list(paging = FALSE, searching = FALSE, processing = FALSE, initComplete = JS(js), preDrawCallback = JS('function() { Shiny.unbindAll(this.api().table().node()); }'), drawCallback = JS('function() { Shiny.bindAll(this.api().table().node()); }')))
    }, server = TRUE)

    # The proxy used for reactive the data frame after user input
    proxy <- dataTableProxy('d2', session = session)

    # Update the table triggered by user edit the cell
    observeEvent(input$d2_cell_edit, {
      info <- input$d2_cell_edit
      i <- info$row
      j <- info$col + 1
      v <- info$value
      result$df[i, j] <- v
      print(result$df)
    })

    # Once click the update, send the row to the SQL database
    observeEvent(input$updated, {
      # Insert row into SQL database
      db_conn <- odbcConnect("ILS")
      sqlQuery(db_conn, "Set Identity_Insert dbo.MarketQuotes On")
      last <- sqlQuery(channel = db_conn, "SELECT TOP 1 id FROM [dbo].[MarketQuotes] ORDER BY id DESC", stringsAsFactors = FALSE)
      id <- last[[1]] + 1
      if (nchar(result$df$CUSIP) > 9) {
        result$df$CUSIP <- substr(result$df$CUSIP, 3, 11)
      }
      manual <- paste("INSERT INTO [dbo].[MarketQuotes]", "(", paste0(SqlName, collapse = ","), ")", "VALUES (", paste("'", result$df$QuoteDate, "',", sep = ""), paste("'", input$BrokerSelect, "',", sep = ""), paste("'", result$df$CUSIP, "',", sep = ""), paste("'", result$df$Bond, "',", sep = ""), paste(result$df$Size, ",", sep = ""), paste("'", input$ActionsSelect, "',", sep = ""), paste(result$df$Price, ",", sep = ""), paste(id, ")", sep = ""), sep = "")
      sqlQuery(db_conn, manual)
    })

    # (Additional functionality like Insert Data, Generate Table, etc. goes here...)

}

# UI definition
ui <- dashboardPage(
  dashboardHeader(title = "Cat Bond Monitor"),
  dashboardSidebar(
    sidebarMenu(
      menuItem("View Table", tabName = "view_table", icon = icon("table")),
      menuItem("Update Table", tabName = "update_table", icon = icon("edit"))
      # Additional menu items here...
    )
  ),
  dashboardBody(
    tabItems(
      tabItem(tabName = "view_table", DTOutput("d1")),
      tabItem(tabName = "update_table", DTOutput("d2"))
      # Additional tab items here...
    )
  )
)

# Run the application
shinyApp(ui = ui, server = server)
