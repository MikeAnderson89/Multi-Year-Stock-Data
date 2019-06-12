Option Explicit

Sub Stocks()
    Dim k As Long
    Dim p As Long
    Dim t As Long
    Dim NumRows As Double
    Dim Total_Ticker_Rows As Long
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Min_Date As Long
    Dim Max_Date As Long
    Dim Starting_Value As Double
    Dim Closing_Value As Double
    Dim Ticker As String
    Dim row As Double
    Dim Symbol As String
    Dim Volume As Double
    Dim sheet As Worksheet
    Dim Ticker_Row As Long
    Dim Min_Change As Double
    Dim Max_Change As Double
    Dim Max_Volume As Double



    For Each sheet In Worksheets

        sheet.Activate

        'defines the number of rows per sheet'
        Range("A1").Select
        Range(Selection, Selection.End(xlDown)).Select
        NumRows = Selection.Rows.Count

        Min_Date = WorksheetFunction.Min(Range("B2:B" & NumRows))
        Max_Date = WorksheetFunction.Max(Range("B2:B" & NumRows))

        'Building the table for the total stock volumes'
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        'Building the table for Greatest % Increase / Decrease'
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"

        'Loops through the tickers for next unique value'
        Symbol = Range("A2").Value
        Ticker_Row = 2
        Volume = 0
        Starting_Value = Range("C2").Value

        For row = 2 To NumRows + 1
            If Cells(row, 1).Value = Symbol Then
                Volume = Volume + Cells(row, 7).Value

                'Closing Value'
                Closing_Value = Cells(row, 6).Value
            Else
                Cells(Ticker_Row, 9).Value = Symbol
                Cells(Ticker_Row, 12).Value = Volume

                'Yearly Change'
                Cells(Ticker_Row, 10).Value = Closing_Value - Starting_Value
                Yearly_Change = Cells(Ticker_Row, 10).Value

                'Percent Change'
                If Starting_Value <> 0 Then
                    Cells(Ticker_Row, 11).Value = Yearly_Change / Starting_Value
                    Percent_Change = Cells(Ticker_Row, 11).Value
                Else
                    Cells(Ticker_Row, 11).Value = "N/A"
                End If

                'Next Ticker'
                Ticker_Row = Ticker_Row + 1
                Symbol = Cells(row, 1).Value
                Volume = Cells(row, 7).Value
                Starting_Value = Cells(row, 6).Value
            End If
        Next row

        'Total Ticker Rows'
        Range("I1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Total_Ticker_Rows = Selection.Rows.Count


        'Finds Min/Max Values for Yearly_Change and Volume'
        Min_Change = WorksheetFunction.Min(Range("K2:K" & Total_Ticker_Rows))
        Max_Change = WorksheetFunction.Max(Range("K2:K" & Total_Ticker_Rows))
        Max_Volume = WorksheetFunction.Max(Range("L2:L" & Total_Ticker_Rows))


        'Loops through Tickers to find values'
        For t = 2 To Total_Ticker_Rows
            If Cells(t, 11).Value = Min_Change Then
                Range("O3").Value = Cells(t, 9).Value
                Range("P3").Value = Cells(t, 11).Value
            ElseIf Cells(t, 11).Value = Max_Change Then
                Range("O2").Value = Cells(t, 9).Value
                Range("P2").Value = Cells(t, 11).Value
            End If
        Next t


        For t = 2 To Total_Ticker_Rows
            If Cells(t, 12).Value = Max_Volume Then
                Range("O4").Value = Cells(t, 9).Value
                Range("P4").Value = Cells(t, 12).Value
            End If
        Next t


        'Formatting'
        sheet.Columns.AutoFit
        Range("K2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.NumberFormat = "0.00%"
        Range("P2").NumberFormat = "0.00%"
        Range("P3").NumberFormat = "0.00%"

        'Conditional Formatting for Yearly Change'
        For t = 2 To Total_Ticker_Rows
            If Cells(t, 10).Value >= 0 Then
                Cells(t, 10).Interior.ColorIndex = 4
            ElseIf Cells(t, 10).Value < 0 Then
                Cells(t, 10).Interior.ColorIndex = 3
            End If
         Next t

    Next sheet

End Sub
