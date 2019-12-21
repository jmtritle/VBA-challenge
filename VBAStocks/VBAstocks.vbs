Attribute VB_Name = "Module1"
Sub stocks():

' Defining i (row), sumrow (summary row) Integer
' Defining ticker (ticker symbol) as String
' Defining eprice (end price), pchange (percent change), sprice (start price), stotal (total stock volume), and ychange (yearly change)
' Defining ws (worksheet number) as Worksheet
' Defining lastrow (dynamic last row)

Dim i, sumrow As Integer
    i = 1
Dim ticker As String
Dim eprice, pchange, sprice, stotal, ychange
Dim ws As Worksheet
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Note on limitations: code only functions correctly if data is sorted alphabetically and chronologically (oldest to newest)

' Initializing the worksheet loop, selecting and activating current worksheet, and setting sumrow to 2
For Each ws In ThisWorkbook.Worksheets
    ws.Select
    ws.Activate
    sumrow = 2

' Setting the summary headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest Percent Increase"
ws.Range("O3").Value = "Greatest Percent Decrease"
ws.Range("O4").Value = "Greatest Total Value"

' Initializing the For Loop
    For i = 2 To lastrow
        ' Setting the start price by determining if previous record is not equal to current
        ' Grabs the record from the first start price available
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                sprice = ws.Cells(i, 3).Value
        End If
        
        ' Setting up the summary with tickers and appropriate calculations
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Setting the ticker symbol
            ticker = ws.Cells(i, 1).Value
                ws.Range("I" & sumrow).Value = ticker
            ' Setting the total stock volume total
            stotal = stotal + ws.Cells(i, 7).Value
                ws.Range("L" & sumrow).Value = stotal
            ' Setting the end price by grabbing the value from the last line of the current ticker
            eprice = ws.Cells(i, 6).Value
            ' Calculating Yearly Change
            ychange = eprice - sprice
                ws.Range("J" & sumrow).Value = ychange
                    If ychange < 0 Then
                        ws.Range("J" & sumrow).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & sumrow).Interior.ColorIndex = 4
                    End If
            ' Calculating Percent Change - if start price is 0, percent change is set to 1 (for 100%)
                    If sprice = 0 Then
                        pchange = 1
                    Else
                        pchange = (eprice - sprice) / sprice
                    End If
                ws.Range("K" & sumrow).Value = pchange
                ws.Range("K" & sumrow).NumberFormat = "0.00%"
            ' Resetting total stock volume, end price, and start price to 0
            stotal = 0
            eprice = 0
            sprice = 0
            ' Incrementing summary row to set up for next record
            sumrow = sumrow + 1
        Else
            ' Adding to total stock volume
            stotal = stotal + ws.Cells(i, 7).Value

        End If

    Next i
    
    ' Defining prange and tsvrange as ranges for WorksheetFunction.Max and .Min
    Dim prange, tsvrange As Range
    ' Setting prange (percent change range) and tsvrange (total stock value range) to pull from the appropriate column
    Set prange = Range("K:K")
    Set tsvrange = Range("L:L")

    ' Using WorksheetFunction.Max to determine greatest percent increase
    Range("Q2") = Application.WorksheetFunction.Max(prange)
    ' Setting the number format to percentage
    Range("Q2").NumberFormat = "0.00%"
    ' Using index and match to find the ticker from column I for the greatest percent increase in Q2 by matching in column K
    Range("P2") = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Range("Q2").Value, Range("K:K"), 0))

    ' Using WorksheetFunction.Min to determine greatest percent decrease
    Range("Q3") = Application.WorksheetFunction.Min(prange)
    ' Setting the number format to percentage
    Range("Q3").NumberFormat = "0.00%"
    ' Using index and match to find the ticker from column I for the greatest percent decrease in Q2 by matching in column K
    Range("P3") = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Range("Q3").Value, Range("K:K"), 0))

    ' Using WorksheetFunction.Max to determine highest total stock volume
    Range("Q4") = Application.WorksheetFunction.Max(tsvrange)
    ' Using index and match to find the ticker from column I for the highest total stock volume in Q2 by matching in column L
    Range("P4") = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Range("Q4").Value, Range("L:L"), 0))
    
  Next ws
  
End Sub

