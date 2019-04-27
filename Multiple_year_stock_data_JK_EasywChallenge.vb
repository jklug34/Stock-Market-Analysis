'# Unit 2 | Assignment - The VBA of Wall Street
'
'## Background
'
'You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, choose your assignment from Easy, Moderate, or Hard below.
'
'### Files
'
'* [Test Data](Resources/alphabtical_testing.xlsx) - Use this while developing your scripts.
'
'* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.
'
'### Stock market analyst
'
'![stock Market](Images/stockmarket.jpg)
'
'### Easy
'
'* Create a script that will loop through one year of stock data for each run and return the total volume each stock had over that year.
'
'* You will also need to display the ticker symbol to coincide with the total stock volume.
'
'* Your result should look as follows (note: all solution images are for 2015 data).
'
'![easy_solution](Images/easy_solution.png)
'

Sub StockMarketEasy():

'Define variables
Dim Ticker As String
Dim Total_Volume As Long
Dim i As Long
Dim Volume As Long
Dim LastRow As Long
Dim Summary_Table_Row As Long
Dim WorksheetName As String
Dim ws As Worksheet


'Dim OpenValue As Double
'Dim ClosingValue As Double
'Dim Yearly_Change As Double
'Dim Percent_Change As Double
'Dim StockDate As Long
'Dim High As Double
'Dim Low As Double


'Initalize the varible "Total Stock Volume" and set to zero
Total_Volume = 0

Summary_Table_Row = 2

        
    For Each ws In Worksheets
        ws.Activate
        Debug.Print ws.Name
        
        ' Add the word Ticker to the Ninth Column Header
        ws.Cells(1, 9).Value = "Ticker"
            
        ' Add the word Total Stock Volume to the Tenth Column Header
        ws.Cells(1, 10).Value = "Total Stock Volume"
            
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
                                    
                'Add row ticker volume to the final stock ticker volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                                    
                'Print the Ticker symbol in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                                    
                'Print the total volume for that ticker symbol in the summary table
                ws.Range("J" & Summary_Table_Row).Value = Total_Volume
                                    
                'Add one to the summary table row for each new ticker
                Summary_Table_Row = Summary_Table_Row + 1
                                    
                'Reset the ticker symbol total volume for next ticker symbol
                Total_Volume = 0
                        
            'If the ticker symbol in the next row is the same then..
            Else
                    
                'Add the ticker volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                'Reset the ticker symbol total volume for next ticker symbol
                Total_Volume = 0
                 
                
            End If
            
                
                  
        Next i
            
                        
            'Reset the Summary_Table_Row back to 2 for the next worksheet
            Summary_Table_Row = 2
        
                             
    Next
    
    
End Sub
