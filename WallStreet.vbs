Sub WallStreet()
'Hard Attempt/Challenge calling worksheet
Dim Ticker As String
Dim TotalVolume As Double
Dim RowNum As Integer
Dim OpenPrice As Double
Dim ClosingPrice As Double
Dim ws As Worksheet

For Each ws In Worksheets

ws.Range("J1") = "Ticker"
ws.Range("K1") = "Yearly Price Change"
ws.Range("L1") = "Percent Change"
ws.Range("M1") = "Total Volume"
ws.Range("Q1") = "Ticker"
ws.Range("R1") = "Value"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  
    
TotalVolume = 0
RowNum = 2
OpenPrice = ws.Cells(2, 3).Value

    
    For r = 2 To LastRow
    
         If (ws.Cells(r + 1, 1).Value) <> (ws.Cells(r, 1).Value) Then
    
            Ticker = ws.Cells(r, 1).Value
            
            ClosingPrice = ws.Cells(r, 6).Value

            TotalVolume = TotalVolume + ws.Cells(r, 7).Value
            
            ws.Range("J" & RowNum).Value = Ticker
   
            ws.Range("K" & RowNum).Value = ClosingPrice - OpenPrice
            
            If (ClosingPrice - OpenPrice) < 0 Then
            
                ws.Range("K" & RowNum).Interior.ColorIndex = 3
            Else
                ws.Range("K" & RowNum).Interior.ColorIndex = 4
            End If
            
          'Check to see if the denominator = 0
            If OpenPrice = 0 Then
                ws.Range("L" & RowNum).Value = 0
           Else
                ws.Range("L" & RowNum).Value = FormatPercent(((ClosingPrice - OpenPrice) / OpenPrice))
           End If
              
            ws.Range("M" & RowNum).Value = TotalVolume
   
            RowNum = RowNum + 1
            TotalVolume = 0
            OpenPrice = ws.Cells(r + 1, 3).Value
            ClosingPrice = 0
   
    Else
            TotalVolume = TotalVolume + ws.Cells(r, 7).Value
                    
     End If
  
  Next r
  
  'Locate the stock with greatest % Increase, greatest % decrease and the greatest total volume
  
  Dim BestPerformer As Double
  Dim WorstPerformer As Double
  Dim GreatestTotalVolume As Double
  Dim LastColunmRow As Integer

  
  BestPerformer = ws.Cells(2, 12).Value
  WorstPerformer = ws.Cells(2, 12).Value
  GreatestTotalVolume = ws.Cells(2, 13).Value
  LastColunmRow = ws.Range("L" & Rows.Count).End(xlUp).Row
    
  
  For r = 2 To LastColunmRow
  
    If ws.Cells(r + 1, 12) > BestPerformer Then
        
        BestPerformer = ws.Cells(r + 1, 12).Value
        
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("P2") = ws.Cells(r + 1, 10).Value
        ws.Range("Q2") = ws.Cells(r + 1, 12).Value
        'ws.Range("Q2").NumberFormat = "0.00%"
       
    End If
    
    If ws.Cells(r + 1, 12) < WorstPerformer Then
        
        WorstPerformer = ws.Cells(r + 1, 12).Value
        
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("P3") = ws.Cells(r + 1, 10).Value
        ws.Range("Q3") = ws.Cells(r + 1, 12).Value
        'ws.Range("Q3").NumberFormat = "0.00%"
    End If
    
        If ws.Cells(r + 1, 13) > GreatestTotalVolume Then
        
        GreatestTotalVolume = ws.Cells(r + 1, 13).Value
        
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P4") = ws.Cells(r + 1, 10).Value
        ws.Range("Q4") = ws.Cells(r + 1, 13).Value
        
        End If
    
Next r
    
Next
 
End Sub







