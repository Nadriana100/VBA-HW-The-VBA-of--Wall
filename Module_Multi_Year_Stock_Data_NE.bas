Attribute VB_Name = "Module1"
' THE MAGIC ACTING IN EVERY WORKSHEET

Sub AllWorksheets()

Dim ws As Worksheet

Application.ScreenUpdating = False

    For Each ws In ActiveWorkbook.Sheets

    ws.Select

    Call The_Wall_Street

    Next ws

Application.ScreenUpdating = True

End Sub

Sub The_Wall_Street()


' 1. The ticker symbol and AND  **4. Total stock volume of the stock

'Group each Ticker symbol
Dim Ticker_Name As String

'Total Stock Volume per ticker
Dim Total_Vol As LongLong

'==========================================

'Summary
Dim Summary As Integer
Summary = 2

'==========================================

'Get the last row
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'===========================================

'2. Yearly change from opening price at the beginning of a given year to_
' the closing price at the end of that year.

'Get Open Vol per ticker symbol
Dim Stk_Open As Double

'Get Close Vol per ticker symbol
Dim Stk_Close As Double


'===========================================

'===========================================


'HERE WE GO!!!!!!!!!!


'Looping the Tickers
    For i = 2 To lastRow

    If Cells(i, 1) <> Cells(i - 1, 1) Then
    Stk_Open = Cells(i, 3).Value
    
    End If

' Grouping ticker's symbols code
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker_Name = Cells(i, 1).Value

' Getting the Yearly Change
Stk_Close = Cells(i, 6).Value
Y_Change = Stk_Close - Stk_Open

' Storing the Yearly Change
Range("J" & Summary).Value = Y_Change

' Getting the Percent Change
'3. The percent change from opening price at the begining of a given year to_
'the closing price at the end of that year.

    If Stk_Open <> 0 Then 'For zeros in close and open
    
Range("K" & Summary).Value = (Stk_Close - Stk_Open) / Stk_Open
Range("K" & Summary).Style = "Percent"

    Else
    Range("K" & Summary).Value = 0
    End If

' Adding the stock volume
Total_Vol = Total_Vol + Cells(i, 7).Value

' Storing the whole tickers symbols
Range("I" & Summary).Value = Ticker_Name
' Storing ticker's total volume
Range("L" & Summary).Value = Total_Vol


Summary = Summary + 1

Total_Vol = 0

Else

'Grouping the same ticker symbols
Total_Vol = Total_Vol + Cells(i, 7).Value

        End If
   
'===========================================
'5. Highlight positive changes in green and negative changes in red.
   
  
    If Cells(i, 10) >= 0 Then
Cells(i, 10).Interior.ColorIndex = 4
    Cells(i, 10).Font.ColorIndex = 1
    Else
Cells(i, 10).Interior.ColorIndex = 3
    Cells(i, 10).Font.ColorIndex = 1
    End If
    
        
Next i

End Sub



'+++++++++++++++++++++++++++++++++++++++++++++++++++++++
'THE END


