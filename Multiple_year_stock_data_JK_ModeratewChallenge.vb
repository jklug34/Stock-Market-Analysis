'### Moderate
'
'* Create a script that will loop through all the stocks for one year for each run and take the following information.
'
'  * The ticker symbol.
'
'  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'
'  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'
'  * The total stock volume of the stock.
'
'* You should also have conditional formatting that will highlight positive change in green and negative change in red.
'
'* The result should look as follows.
'
'![moderate_solution](Images/moderate_solution.png)


Sub StockMarketModerate():

'Define variables
Dim Ticker As String
Dim Total_Volume As Double
Dim i As Long
Dim Volume As Long
Dim LastRow As Long
Dim Summary_Table_Row As Integer
Dim WorksheetName As String
Dim ws As Worksheet
Dim OpeningValue As Double
Dim ClosingValue As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim OpeningValueAll() As Variant




'Dim OpenValue As Double
'Dim ClosingValue As Double
'Dim Yearly_Change As Double
'Dim Percent_Change As Double
'Dim StockDate As Long
'Dim High As Double
'Dim Low As Double


'Initalize the varible "Total Stock Volume" and set to zero
Total_Volume = 0
OpeningValue = 0
ClosingValue = 0
Yearly_Change = 0
Percent_Change = 0

Summary_Table_Row = 2

        
    For Each ws In Worksheets
        ws.Activate
        Debug.Print ws.Name
        
        ' Add the word Ticker to the Ninth Column Header
        ws.Cells(1, 9).Value = "Ticker"
               
        ' Add the word Yearly Change to the Tenth Column Header
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ' Add the word Yearly Change to the Tenth Column Header
        ws.Cells(1, 11).Value = "Percent Change"
        
        ' Add the word Total Stock Volume to the Tenth Column Header
        ws.Cells(1, 12).Value = "Total Stock Volume"
            
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
            
        
        ' Loop through all the ticker values
        For i = 2 To LastRow
            
                
            'Check if we are still within the same ticker symbol, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then '<> is not equal to, cell (3,2) is not equal to cell (2,2) and so on...
                            
                'Set the ticker symbol
                Ticker = ws.Cells(i, 1).Value
                
                'Grab the Opening Value for each ticker
                ClosingValue = ws.Cells(i, 6).Value
                'Debug.Print ClosingValue
                                        
                'Add row ticker volume to the final stock ticker volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                                
                'Print the Ticker symbol in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                        
                'Print the Yearly_Change in the summary table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                        
                'PrBCint the Percent_Change in the summary table
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                        
                'Print the total volume for that ticker symbol in the summary table
                ws.Range("L" & Summary_Table_Row).Value = Total_Volume
                        
                'Add one to the summary table row for each new ticker
                Summary_Table_Row = Summary_Table_Row + 1
                                    
                'Reset the ticker symbol total volume for next ticker symbol
                Total_Volume = 0
                OpeningValue = 0
                ClosingValue = 0
                Yearly_Change = 0
                Percent_Change = 0
            
                        
            'If the ticker symbol in the next row is the same then..
            Else
                        
                'Add the ticker volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                
            End If
            
                'Grab the Opening Value for each ticker
                OpeningValueAll = Array(ws.Cells(i, 3).Value)
                OpeningValue = OpeningValueAll(1) 'find the first element of the array
                'Debug.Print OpeningValue
                
                'Define the Yearly Change in Value as the Closing Value - the Opening Value
                Yearly_Change = ClosingValue - OpeningValue
                
                'Define the Percent Change as the Yearly_Change - the Opening Value
                Percent_Change = (Yearly_Change / OpeningValue) * 100
                  
        Next i
                        
            'Reset the Summary_Table_Row back to 2 for the next worksheet
            Summary_Table_Row = 2
        
        '        Set as Percent
                '-------------------------------------------------------
        
                 ws.Range("K2:K" & LastRow).Style = "Percent"
        '
                 'OR
        '
        '        ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-8]"
        '        Selection.NumberFormat = "0.00%"
        '
        '        Set the conditional formatting rules
                 '-----------------------------------------------------
        '
                 If Percent_Change > 0 Then
        '
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        '
                 Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
                 End If
                 
                 
                 'OR
        '
        '            Columns("J:J").Select
        '        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        '            Formula1:="=0"
        '        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        '        With Selection.FormatConditions(1).Interior
        '            .PatternColorIndex = xlAutomatic
        '            .Color = 5296274
        '            .TintAndShade = 0
        '        End With
        '        Selection.FormatConditions(1).StopIfTrue = False
        '        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        '            Formula1:="=0"
        '        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        '        With Selection.FormatConditions(1).Interior
        '            .PatternColorIndex = xlAutomatic
        '            .Color = 255
        '            .TintAndShade = 0
        '        End With
        '        Selection.FormatConditions(1).StopIfTrue = False
                        
    Next

End Sub


