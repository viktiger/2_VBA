Attribute VB_Name = "Module1"
Sub MakeMoneyPart1a()

'---PART 1---Complete Task for all worksheets

'Dim ws As Worksheet

'For Each ws In ThisWorkbook.Worksheets


' Setup Ticker Symbols in column L ensuring no duplicates
    ' Copy Entire Column A
    Sheets("2018").Select
    Columns("A:A").Select
    Selection.Copy
    
    ' Paste in Column L and remove duplicates
    Columns("L:L").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("$L$1:$L$760000").RemoveDuplicates Columns:=1

' Naming columns L to Q with header in quotation marks e.g. "Ticker"
    Range("L1").Value = "Ticker"
    Range("M1").Value = "Opening Price"
    Range("N1").Value = "Closing Price"
    Range("O1").Value = "Annual Change"
    Range("P1").Value = "Annual Percentage Change"
    Range("Q1").Value = "Share Volume"

'---PART 2 PREWORK---
' Determine Opening Price
    ' row count for dataset or column L
    For j = 2 To 3100

    ' row count for output table or column A
        For i = 2 To 760000
            If Cells(i, 1).Value = Cells(j, 12).Value Then
                Cells(j, 13).Value = Cells(i, 3).Value
            Exit For
            End If
        Next i
    Next j

'Next ws

End Sub

Sub MakeMoneyPart1b()

'---PART 2 PREWORK---Complete Task for all worksheets

'Dim ws As Worksheet

'For Each ws In ThisWorkbook.Worksheets

' Determine Closing Price
    ' row count for dataset or column A
    For i = 2 To 760000

    ' row count for output table or column L
        For j = 2 To 3100
            If ws.Cells(j, 12).Value = ws.Cells(i, 1).Value Then
                ws.Cells(j, 14).Value = ws.Cells(i, 6).Value
            End If
        Next j
    Next i

'Next ws

End Sub

Sub MakeMoneyPart2()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

' Naming columns L to Q with header in quotation marks e.g. "Ticker"
    ws.Range("L1").Value = "Ticker"
    ws.Range("M1").Value = "Opening Price"
    ws.Range("N1").Value = "Closing Price"
    ws.Range("O1").Value = "Annual Change"
    ws.Range("P1").Value = "Annual Percentage Change"
    ws.Range("Q1").Value = "Share Volume"

' Formatting Opening, Closing Price, Annual Change, Percentage Change, Share Volume Fields
    ws.Range("M2:M3100").NumberFormat = "#,##0.00_);(#,##0.00);?_)"
    ws.Range("N2:N3100").NumberFormat = "#,##0.00_);(#,##0.00);?_)"
    ws.Range("O2:O3100").NumberFormat = "#,##0.00_);(#,##0.00);?_)"
    ws.Range("P2:P3100").NumberFormat = "#,##0.00%_);(#,##0.00%);?_)"
    ws.Range("Q2:Q3100").NumberFormat = "#,##0_);(#,##0);?_)"

'---PART 2---
' Calculating the Yearly Change
    ' For every row Closing Price minus Opening Price
    For x = 2 To 3100
        ws.Cells(x, 15).Value = ws.Cells(x, 14).Value - ws.Cells(x, 13).Value
    Next

'---PART 3---
' Calculating the Percentage Change
    ' For every row Yearly Change divided by Opening Price except if Opening Price is zero then Yearly Change is zero
    For y = 2 To 3100
        If ws.Cells(y, 13).Value = 0 Then
            ws.Cells(y, 16).Value = 0

        Else: ws.Cells(y, 16).Value = (ws.Cells(y, 15).Value / ws.Cells(y, 13).Value)
        End If
    Next

'---PART 4---
' Calculating the Stock Volume
    ws.Range("Q2:Q3100").FormulaR1C1 = "=SUMIFS(C[-10],C[-16],RC[-5])"

Next ws

End Sub

Sub MakeMoneyBonus()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

ws.Range("S1").Value = "Bonus"

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotalVolume As Double

GreatestIncrease = 0
GreatestDecrease = 0
GreatestTotalVolume = 0

' Inserting calculation descriptions in cells "S2","S3","S4"
ws.Range("S2").Value = "Greatest % Increase"
ws.Range("S3").Value = "Greatest % Decrease"
ws.Range("S4").Value = "Greatest Total Stock Volume"

' Formatting Percentage Increase & Decrease and Total Stock Volume Fields
ws.Range("T2:T3").NumberFormat = "#,##0.00%_);(#,##0.00%);?_)"
ws.Range("T4:T6").NumberFormat = "#,##0_);(#,##0);?_)"

' Identify Greatest % Increase
For i = 2 To 3100
    If ws.Cells(i, 16).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(i, 16).Value
        ws.Range("T2").Value = GreatestIncrease
        End If
    Next i
    
' Identify Greatest % Decrease
For j = 2 To 3100
    If ws.Cells(j, 16).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(j, 16).Value
        ws.Range("T3").Value = GreatestDecrease
        End If
    Next j

' Identify Greatest Total Volume
For k = 2 To 3100
    If ws.Cells(k, 17).Value > GreatestTotalVolume Then
        GreatestTotalVolume = ws.Cells(k, 17).Value
        ws.Range("T4").Value = GreatestTotalVolume
        End If
    Next k

' Format background color grey to visualise the bonus section
' Range("S2:T4").Interior.ColorIndex = 15

Next ws

End Sub

