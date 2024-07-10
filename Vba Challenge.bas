Attribute VB_Name = "Module1"
Sub Answer():

'Assigning Variables
Dim i As Long
Dim Ticker As String
Dim lastRow As Long
Dim SummaryTableRow As Integer
Dim TotalStockVolume As LongLong
Dim QuarterlyChange As Double
Dim FirstRow As Boolean
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim ws As Worksheet

'Loop through worksheets
For Each ws In ThisWorkbook.Worksheets

    'Variable Values
    FirstRow = True
    TotalStockVolume = 0
    SummaryTableRow = 2
    QuarterlyChange = 0
    
    
    'Defining Last Row of Data
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
  
    'For Column Titles
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    
    'Start of Loop through Stock Data
        For i = 2 To lastRow
        
    'Capturing the Stocks Opening Price
            If FirstRow Then
                OpenPrice = ws.Cells(i, 3).Value
                FirstRow = False
        
    'Finding where the Ticker switches to a different stock
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
    'Print Ticker name, Percent Change, and Quarter Change in the Summary Table
                Ticker = ws.Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                ClosePrice = ws.Cells(i, 6).Value
                QuarterlyChange = ClosePrice - OpenPrice
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                ws.Range("J" & SummaryTableRow).Value = QuarterlyChange
                ws.Range("K" & SummaryTableRow).Value = (QuarterlyChange / OpenPrice)
                
    'Advancing Summary Table down one row for next Stock
                SummaryTableRow = SummaryTableRow + 1
              
    'Reseting values for next stock
                OpenPrice = ws.Cells(i + 1, 3).Value
                QuarterlyChange = 0
                TotalStockVolume = 0
                
    'For non first row and non last row this keeps the stock volume addition going
            Else
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                        
            End If
              
            
        Next i
        
     
        
    'Looping back through the data for formatting
    For i = 2 To lastRow
    
        If ws.Cells(i, 10).Value > "0" Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < "0" Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        
    'Formatting for columns to fit the data provided
    Next i
     ws.Columns("J:J").EntireColumn.AutoFit
     ws.Columns("L:L").EntireColumn.AutoFit
     ws.Columns("K:K").EntireColumn.AutoFit
     ws.Columns("I:I").EntireColumn.AutoFit
     
 
 Next ws
 
 'Formatting ColumnK to be a Percent
 For Each ws In ThisWorkbook.Worksheets
        
        Set col = ws.Columns("K")
        
        
        col.NumberFormat = "0.00%"
        
    
Next ws

'Variables for finding the Greatest increase and decrease
Dim lastRowTotal As Long
Dim MaxChange As Double
Dim CurrentValue As Double


For Each ws In Worksheets

'Cell titles
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
      ' Defining LastRow and MaxChange Values
        lastRowTotal = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        
        MaxChange = ws.Cells(2, 11).Value
        
        ' Loop to find Greatest Percent Increase
        
        For i = 2 To lastRowTotal
            CurrentValue = ws.Cells(i, 11).Value
           
            If CurrentValue >= MaxChange Then
                MaxChange = CurrentValue
                ws.Cells(2, 16).Value = MaxChange
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                
            End If
            
        Next i
        
        'Loop to find Greatest Percent Decrease
        MaxChange = ws.Cells(2, 11).Value
        For i = 2 To lastRowTotal
            CurrentValue = ws.Cells(i, 11).Value
           
            If CurrentValue <= MaxChange Then
                MaxChange = CurrentValue
                ws.Cells(3, 16).Value = MaxChange
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                
            End If
            
        Next i
         
        'Loop to find Greatest Volume
         MaxChange = ws.Cells(2, 12).Value
        For i = 2 To lastRowTotal
            CurrentValue = ws.Cells(i, 12).Value
           
            If CurrentValue >= MaxChange Then
                MaxChange = CurrentValue
                ws.Cells(4, 16).Value = MaxChange
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                
            End If
         
        'Formatting columns to fit
         Next i
        ws.Columns("N:N").EntireColumn.AutoFit
        ws.Columns("O:O").EntireColumn.AutoFit
        ws.Columns("P:P").EntireColumn.AutoFit
        
Next ws

'Format Percentage in summary table
For Each ws In ThisWorkbook.Worksheets
        With ws
           
            .Range("P2").NumberFormat = "0.00%"
            .Range("P3").NumberFormat = "0.00%"
        End With
    Next ws

 
'Fin
End Sub
