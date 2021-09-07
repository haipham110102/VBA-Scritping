Attribute VB_Name = "Module1"
Sub StockData()

  Dim ws As Worksheet
  For Each ws In Worksheets
        Dim WorksheetName As String
            WorksheetName = ws.Name
            'MsgBox (WorksheetName)
        
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim Ticker_Name As String
        Dim TotalTickerVolume As Double
            TotalTickerVolume = 0
    
        Dim Summary_Table_Row As Long
            Summary_Table_Row = 2
        
        Dim OpenValue As Long
            OpenValue = 2
        
        Dim LastRowTicker As Long
        
            'Column Headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            'Bonus Part: Name and Column Headers
        Dim GreatestIncrease As Double
            GreatestIncrease = 0
        Dim GreatestDecrease As Double
            GreatestDecrease = 0
        Dim GreatestTotalVolume As Double
            GreatestTotalVolume = 0
        Dim LastRowPercentChange As Long
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            
    
        LastRowTicker = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        For i = 2 To LastRowTicker
            

            'Check if we are still within the same ticker, if it is not
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Add ticker Total Volume
            TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
                        
            'Set the ticker Name
            Ticker_Name = ws.Cells(i, 1).Value
            

            'Print the ticker Name in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
        
            'Print the ticker volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = TotalTickerVolume
            
            'Reset the Sticker Total
            TotalTickerVolume = 0
            
         
               
            'get OpenPrice
            OpenPrice = ws.Range("C" & OpenValue)
            
            'print (testing) Openprice to S column before calculate percentage
            'ws.Range("S" & Summary_Table_Row).Value = OpenPrice
            
            'get ClosePrice
            ClosePrice = ws.Range("F" & i)
            
            'print (Testing) Openprice to S column before calculate percentage
            'ws.Range("T" & Summary_Table_Row).Value = ClosePrice
            
            'Calculate Yearly Change
            YearlyChange = ClosePrice - OpenPrice
            
            'Print YearlyChange to Summary Table
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            
            'formatting Yearly Change:$0.00
            ws.Range("J" & Summary_Table_Row).NumberFormat = "$0.00"
                
                'Format Yearly change cell color: Green >=0 and Red < 0
                If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                   ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                Else
                   ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
                   
                'Calculate Percent Change and Avoiding #DIV/0! when OpenPrice is 0
                If OpenPrice = 0 Then
                   PercentChange = 0
                
                Else
            
                   YearlyOpen = ws.Range("C" & OpenValue)
                   
                   'Calculate Percent Change
                   PercentChange = YearlyChange / OpenPrice
                   
                   'print to Summary Table
                   ws.Range("K" & Summary_Table_Row).Value = PercentChange
                   
                   'format Percent Change:0.00%
                   ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                   
                End If
                
                
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'get Open Value from each ticker
            OpenValue = i + 1

                   
              
            'If the cell immediately following a row is the same Ticker ...
            Else
            
            'Add to the Total Ticker Volume
            TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
            
            

            End If
        
        Next i
        
    'Bonus Part
    'Determine LastRow of Percent Change (Column K)
    LastRowPercentChange = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        For i = 2 To LastRowPercentChange
    
            'Calculate Greatest percent Increase
    
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
               ws.Range("Q2").Value = ws.Range("K" & i).Value
               ws.Range("P2").Value = ws.Range("I" & i).Value
    
                        
            'Calculate Greatest Percent Decrease
            ElseIf ws.Range("K" & i).Value < ws.Range("Q3").Value Then
               ws.Range("Q3").Value = ws.Range("K" & i).Value
               ws.Range("P3").Value = ws.Range("I" & i).Value
               
            'Formating Greatest percent increase and decrease 0.00%
               ws.Range("Q2:Q3").NumberFormat = "0.00%"
            
            End If
            
            'Calculate Greatest Total Volume
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
               ws.Range("Q4").Value = ws.Range("L" & i).Value
               ws.Range("P4").Value = ws.Range("I" & i).Value
               
            End If
               
            
        Next i
        
  Next ws
  
  MsgBox ("Done")

End Sub
