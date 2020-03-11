Attribute VB_Name = "Module1"
Sub StockAnalysis()
    
    'set ticker counter variable
    Dim Ticker As String
    'set stock volume variable
    Dim volume As Double
        volume = 0
    'set opening price variable
    Dim openprice As Double
    'set closing price variable
    Dim closeprice As Double
    'set yearlychange variable
    Dim yearlychange As Double
    'set percentchange variable
    Dim percentchange As Double
    'set up summary row placement
    Dim summarytablerow As Integer
    'set up variable to review worksheets
    Dim currentws As Worksheet
    For Each currentws In Worksheets
    
    ' Make the table headers
    'This one holds the ticker info
    currentws.Cells(1, 9).Value = "Ticker"
    'This one holds the yearly change info
    currentws.Cells(1, 10).Value = "Yearly Change"
    'This one holds the percent change info
    currentws.Cells(1, 11).Value = "Percent Change"
    'This one holds the stock volume info
    currentws.Cells(1, 12).Value = "Stock Volume"
    
    'This is the first open price
    openprice = currentws.Cells(2, 3).Value
    
    'go through the applicable rows
    lastrow = currentws.Cells(Rows.Count, 1).End(xlUp).Row
    summarytablerow = 2
        'Loop
        For Row = 2 To lastrow
        
        'look to see if the cells are differerent from each other, if they are then..
        If currentws.Cells(Row + 1, 1).Value <> currentws.Cells(Row, 1).Value Then
        
        'store the ticker
        Ticker = currentws.Cells(Row, 1).Value
        
        'and store the stock volume
        volume = volume + currentws.Cells(Row, 7).Value
        
        'update the excel sheet with the ticker and volume
        currentws.Range("I" & summarytablerow).Value = Ticker
        currentws.Range("L" & summarytablerow).Value = volume
        
        'now we want to update the sheet with the yearly change
        'find the closing price to compare
        closeprice = currentws.Cells(Row, 6).Value
        
        'find the yearly change
        yearlychange = (closeprice - openprice)
        
        'Update the excel sheet with the yearlychange amount
        currentws.Range("J" & summarytablerow).Value = yearlychange
        
            'because you receive a divide by zero error...
            'set up the open price
            If (openprice = 0) Then
        
            percentchange = 0
        
            'otherwise
            Else
        
            'calculate the percent change for the stock
            percentchange = yearlychange / openprice
        
            'End this loop
            End If
        
        'Update the excel sheet and calculate it into the required percent
        currentws.Range("K" & summarytablerow).Value = percentchange
        currentws.Range("K" & summarytablerow).NumberFormat = "0.00%"
        
        'move to the next row now
        
        summarytablerow = summarytablerow + 1
        
        'bring the volume back to 0
        volume = 0
        
        'Resetting the open price
        openprice = currentws.Cells(Row + 1, 3)
        
        Else
        
        'we need to add the volume
        volume = volume + currentws.Cells(Row, 7).Value
        
        End If
    Next Row

    lastsummaryrow = currentws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'set up the conditional formatting requirements
    'Loop
    For Row = 2 To lastsummaryrow
    'If it is greater than 0
    If currentws.Cells(Row, 10).Value > 0 Then
       currentws.Cells(Row, 10).Interior.ColorIndex = 10
   'Otherwise, if not
    Else
    
    currentws.Cells(Row, 10).Interior.ColorIndex = 3
    
    End If
    

Next Row

Next currentws

End Sub
