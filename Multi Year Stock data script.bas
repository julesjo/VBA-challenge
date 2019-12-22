Attribute VB_Name = "Module1"
Sub StockTracker()
WC = ActiveWorkbook.Worksheets.Count

For j = 1 To WC

        ActiveWorkbook.Worksheets(j).Range("K1").Value = "Yearly Change"
        ActiveWorkbook.Worksheets(j).Range("J1").Value = "Ticker"
        ActiveWorkbook.Worksheets(j).Range("L1").Value = "Percentage Change"
        ActiveWorkbook.Worksheets(j).Range("M1").Value = "Total Stock Volume"
        ActiveWorkbook.Worksheets(j).Range("N1").Value = "Opening Rate"
        ActiveWorkbook.Worksheets(j).Range("O1").Value = "Closing Rate"
        ActiveWorkbook.Worksheets(j).Range("R1").Value = "Ticker"
        ActiveWorkbook.Worksheets(j).Range("S1").Value = "Value"
        ActiveWorkbook.Worksheets(j).Range("Q2").Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(j).Range("Q3").Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(j).Range("Q4").Value = "Greatest Total Volume"
        
        Dim Ticker As String
        Dim TotalStockVolume As Double
        Dim OpeningRate As Double
        Dim ClosingRate As Double
        Dim YearlyChange As Double
        Dim PercentageChange As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestTotal As Double
        Dim Summary_Table_Row As Integer
                
        YearlyChange = 0
        PercentageChange = 0
        TotalStockVolume = 0
        OpeningRate = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotal = 0
        Summary_Table_Row = 2
        
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
            If OpeningRate = 0 Then
               OpeningRate = Range("C" & i)
            End If
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'OpeningRate = Cells(Rows.Count & "C").Value
                'ClosingRate = Range("F").End(xlDown).Value
                'OpeningRate = Index(Cells(i, 3), Match(True, Cells(i, 3) <> "", 0))
                ClosingRate = Cells(i, 6).Value
                YearlyChange = (ClosingRate - OpeningRate)
                ActiveWorkbook.Worksheets(j).Range("K" & Summary_Table_Row).Value = YearlyChange
                PercentageChange = (ClosingRate - OpeningRate) / ClosingRate * 100
                ActiveWorkbook.Worksheets(j).Range("L" & Summary_Table_Row).Value = PercentageChange & "%"
                If YearlyChange < 0 Then
                    'ActiveWorkbook.Worksheets(j).Range("L" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                    ActiveWorkbook.Worksheets(j).Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    ActiveWorkbook.Worksheets(j).Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
                
                Ticker = Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                ActiveWorkbook.Worksheets(j).Range("J" & Summary_Table_Row).Value = Ticker
                ActiveWorkbook.Worksheets(j).Range("M" & Summary_Table_Row).Value = TotalStockVolume
                ActiveWorkbook.Worksheets(j).Range("N" & Summary_Table_Row).Value = OpeningRate
                ActiveWorkbook.Worksheets(j).Range("O" & Summary_Table_Row).Value = ClosingRate
                
                If GreatestIncrease = 0 Then
                    GreatestIncrease = PercentageChange
                    ActiveWorkbook.Worksheets(j).Cells(2, "S").Value = GreatestIncrease
                    ActiveWorkbook.Worksheets(j).Cells(2, "R").Value = Ticker
                ElseIf PercentageChange > GreatestIncrease Then
                    GreatestIncrease = PercentageChange
                    ActiveWorkbook.Worksheets(j).Cells(2, "S").Value = GreatestIncrease
                    ActiveWorkbook.Worksheets(j).Cells(2, "R").Value = Ticker
                End If
                
                If GreatestDecrease = 0 Then
                    GreatestDecrease = PercentageChange
                    ActiveWorkbook.Worksheets(j).Cells(3, "S").Value = GreatestDecrease
                    ActiveWorkbook.Worksheets(j).Cells(3, "R").Value = Ticker
                ElseIf PercentageChange < GreatestDecrease Then
                    GreatestDecrease = PercentageChange
                    ActiveWorkbook.Worksheets(j).Cells(3, "S").Value = GreatestDecrease
                    ActiveWorkbook.Worksheets(j).Cells(3, "R").Value = Ticker
                End If
                
                If GreatestTotal = 0 Then
                    GreatestTotal = TotalStockVolume
                    ActiveWorkbook.Worksheets(j).Cells(4, "S").Value = GreatestTotal
                    ActiveWorkbook.Worksheets(j).Cells(4, "R").Value = Ticker
                ElseIf TotalStockVolume > GreatestTotal Then
                    GreatestTotal = TotalStockVolume
                    ActiveWorkbook.Worksheets(j).Cells(4, "S").Value = GreatestTotal
                    ActiveWorkbook.Worksheets(j).Cells(4, "R").Value = Ticker
                End If
                
                Summary_Table_Row = Summary_Table_Row + 1
                YearlyChange = 0
                TotalStockVolume = 0
                OpeningRate = 0
                ClosingRate = 0
            Else
                TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                YearlyChange = (ClosingRate - OpeningRate)
            End If
        Next i
    Next j
End Sub

