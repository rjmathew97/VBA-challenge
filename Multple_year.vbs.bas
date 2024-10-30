Attribute VB_Name = "Module1"
Sub Multiyear()

    ' Create variables for  Worksheet
    Dim ws As Worksheet
    Dim ticker As String
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim quarterStart As Double
    Dim quarterEnd As Double
    Dim lastRow As Long
    Dim i As Long
    Dim resultRow As Long

    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    ' Initialize greatest values
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets

        ' Set initial variables for the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        resultRow = 2
        totalVolume = 0

        ' Add headers for the Ticker, Quaterly Change, Percent Change, Total Volume
        Range("T1").Value = "Ticker"
        Range("J1").Value = "Quarterly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Volume"
        
        ' Loop through all rows in the worksheet
        For i = 2 To lastRow

            ' Check if we are at the start of a new ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                ' Calculate quarterly change and percentage change for the previous ticker
                If i > 2 Then
                    quarterEnd = ws.Cells(i - 1, 6).Value ' Closing price of the previous quarter
                    quarterlyChange = quarterEnd - quarterStart
                    
                    ' Avoid division by zero
                    If quarterStart <> 0 Then
                        percentChange = (quarterlyChange / quarterStart) * 100
                    Else
                        percentChange = 0
                    End If
                    
                    ' Output the results to the worksheet
                    ws.Cells(resultRow, 9).Value = ticker
                    ws.Cells(resultRow, 10).Value = quarterlyChange
                    ws.Cells(resultRow, 11).Value = percentChange
                    ws.Cells(resultRow, 12).Value = totalVolume
                    
                    ' Update greatest increase/decrease/volume
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        greatestIncreaseTicker = ticker
                    End If

                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        greatestDecreaseTicker = ticker
                    End If

                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = ticker
                    End If

                    ' Increment the result row
                    resultRow = resultRow + 1
                End If
                
                ' Set variables for the new ticker
                ticker = ws.Cells(i, 1).Value
                quarterStart = ws.Cells(i, 3).Value ' Opening price of the new quarter
                totalVolume = 0
            End If

            ' Add to the total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        Next i

        ' Calculate and output for the last ticker
        quarterEnd = ws.Cells(lastRow, 6).Value ' Closing price of the last quarter
        quarterlyChange = quarterEnd - quarterStart

        If quarterStart <> 0 Then
            percentChange = (quarterlyChange / quarterStart) * 100
        Else
            percentChange = 0
        End If

        ' Output the final results to the worksheet
        ws.Cells(resultRow, 9).Value = ticker
        ws.Cells(resultRow, 10).Value = quarterlyChange
        ws.Cells(resultRow, 11).Value = percentChange
        ws.Cells(resultRow, 12).Value = totalVolume

        ' Update greatest increase/decrease/volume for the last ticker
        If percentChange > greatestIncrease Then
            greatestIncrease = percentChange
            greatestIncreaseTicker = ticker
        End If

        If percentChange < greatestDecrease Then
            greatestDecrease = percentChange
            greatestDecreaseTicker = ticker
        End If

        If totalVolume > greatestVolume Then
            greatestVolume = totalVolume
            greatestVolumeTicker = ticker
        End If
        
        ' Write results to the current worksheet
        With ws
            .Cells(2, 15).Value = "Greatest % Increase"
            .Cells(2, 16).Value = greatestIncreaseTicker
            .Cells(2, 17).Value = greatestIncrease / 100

            .Cells(3, 15).Value = "Greatest % Decrease"
            .Cells(3, 16).Value = greatestDecreaseTicker
            .Cells(3, 17).Value = greatestDecrease / 100

            .Cells(4, 15).Value = "Greatest Total Volume"
            .Cells(4, 16).Value = greatestVolumeTicker
            .Cells(4, 17).Value = greatestVolume
        End With
    Next ws
    
End Sub


