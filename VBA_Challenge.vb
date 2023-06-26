VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockChange()

Dim lastRow As Long
Dim WorksheetName As String
Dim Ticker As String
Dim nextTicker As String
Dim previousTicker As String
Dim lastClosing As Double
Dim firstOpening As Double
Dim results As Integer
Dim Volume_Total As Double
Dim Greatest_Volume As Double


Volume_Total = 0
firstOpening = 0
lastClosing = 0
Greatest_Volume = 0

For Each ws In Worksheets

ws.Range("I1,P1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest Volume"
Range("O3").Value = "Greatest Percent Increase"
Range("O4").Value = "Greatest Percent Decrease"
Range("Q1").Value = "Value"

' Reset greatest and least to 0 (volume)



    ' Setting an initial variable for each ticket for each stock
    Dim Ticker_Row As Long
    
    Ticker_Row = 2
    
    ' Find the last row of column 1
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastRow
        Ticker = ws.Cells(i, 1).Value
        nextTicker = ws.Cells(i + 1, 1).Value
        previousTicker = ws.Cells(i - 1, 1).Value
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value
             ' Check if we are still within the same stock type, if it is not�
             ' Finds the last row that matches; at this last row, it will follow this code:
             If nextTicker <> Ticker Then
               
                
                ' Print the Stock Name in the Ticker Column in Summary Table
                ws.Range("I" & Ticker_Row).Value = Ticker
                
                ' Print the Stock Volume Total to the Summary Table
                ws.Range("L" & Ticker_Row).Value = Volume_Total
                
                lastClosing = Cells(i, 6).Value
                ws.Range("J" & Ticker_Row).Value = lastClosing - firstOpening
                
                    If ws.Range("J" & Ticker_Row).Value > 0 Then
                    ws.Range("J" & Ticker_Row).Interior.ColorIndex = 4
                    
                    ElseIf ws.Range("J" & Ticker_Row).Value < 0 Then
                    ws.Range("J" & Ticker_Row).Interior.ColorIndex = 3
                    
                    End If
                
                ws.Range("K" & Ticker_Row).Value = ws.Range("J" & Ticker_Row).Value / firstOpening
                    
                    If ws.Range("K" & Ticker_Row).Value > 0 Then
                    ws.Range("K" & Ticker_Row).Interior.ColorIndex = 4
                    
                    ElseIf ws.Range("K" & Ticker_Row).Value < 0 Then
                    ws.Range("K" & Ticker_Row).Interior.ColorIndex = 3
                    
                    End If
                
                ' Adding another row to the Summary Table
                Ticker_Row = Ticker_Row + 1
                
                ' Resetting the Volume total so it does not continue building upon one another
                Volume_Total = 0
                
            ' Finds the first row of new stock
            ElseIf previousTicker <> Ticker Then
             
                ' This will give us the firstOpening
                firstOpening = ws.Cells(i, 3).Value
            
                
            End If
            
         
    Next i
    
    ws.Range("Q2") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
    ws.Range("Q3") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
    ws.Range("Q4") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100

    Volume_Number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
    ws.Range("P2") = ws.Cells(Volume_Number + 1, 9)
    
    Increase_Number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    ws.Range("P3") = ws.Cells(Increase_Number + 1, 9)
    
    Decrease_Number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    ws.Range("P4") = ws.Cells(Decrease_Number + 1, 9)
    
    
    
Next ws
    
End Sub

