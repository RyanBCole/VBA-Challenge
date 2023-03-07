Attribute VB_Name = "Module1"
Sub Stock_Data()

    Dim ws As Worksheet
    Dim last_row As Double
    Dim year_opening As Single
    Dim year_closing As Single
    Dim volume As Double
    Dim nextticker As Double
    Dim banana As Double
    
    For Each ws In Sheets
        Worksheets(ws.Name).Activate
        nextticker = 2
        banana = 2
        last_row = WorksheetFunction.CountA(ActiveSheet.Columns(1))
        volume = 0
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
        
        
        'Find unique tickers and add them to new Ticker column
        For i = 2 To last_row
            ticker = Cells(i, 1).Value
            ticker_check = Cells(i + 1, 1).Value
            If ticker <> ticker_check Then
                Cells(banana, 9).Value = ticker
                banana = banana + 1
            End If
         Next i
    
        'Go through all tickers. When ticker changes, start over until done
        For i = 2 To last_row + 1
            ticker = Cells(i, 1).Value
            ticker_check = Cells(i - 1, 1).Value
            If ticker = ticker_check Then
                volume = volume + Cells(i, 7).Value
            ElseIf i > 2 Then
                Cells(nextticker, 12).Value = volume
                nextticker = nextticker + 1
                volume = 0
            Else
                volume = volume + Cells(i, 7).Value
            End If
        Next i
            
        'Go through all tickers. When ticker changes, assign year open and close.
        nextticker = 2
        For i = 2 To last_row + 1
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                year_closing = Cells(i, 6).Value
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                year_opening = Cells(i, 3).Value
            End If
            If year_opening > 0 And year_closing > 0 Then
                increase = year_closing - year_opening
                percent_increase = increase / year_opening
                Cells(nextticker, 10).Value = increase
                Cells(nextticker, 11).Value = FormatPercent(percent_increase)
                year_closing = 0
                year_opening = 0
                nextticker = nextticker + 1
            End If

        Next i
        
        'Find min and max values
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Volume"
        
        max_per = WorksheetFunction.Max(ActiveSheet.Columns("k"))
        min_per = WorksheetFunction.Min(ActiveSheet.Columns("k"))
        max_vol = WorksheetFunction.Max(ActiveSheet.Columns("l"))
        
        Range("Q2").Value = FormatPercent(max_per)
        Range("Q3").Value = FormatPercent(min_per)
        Range("Q4").Value = max_vol
        
        
        'Find min and max % changed
        For i = 2 To last_row
            If max_per = Cells(i, 11).Value Then
                Range("P2").Value = Cells(i, 9).Value
            ElseIf min_per = Cells(i, 11).Value Then
                Range("P3").Value = Cells(i, 9).Value
        'Find max volume and add it to cell
            ElseIf max_vol = Cells(i, 12).Value Then
                Range("P4").Value = Cells(i, 9).Value
            End If
            
        Next i
        
        'Color the change
        For i = 2 To last_row + 1
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
            
        Next i
        
    Next ws
    
End Sub




