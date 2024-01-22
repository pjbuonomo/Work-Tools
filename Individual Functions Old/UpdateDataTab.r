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
