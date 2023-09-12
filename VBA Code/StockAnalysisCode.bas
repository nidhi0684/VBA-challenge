Attribute VB_Name = "Module1"
Sub StockSummary()

    Dim StartRowForParsing As Long
    Dim ColumnForTicker As Integer
    Dim ColumnForOpenPrice As Integer
    Dim ColumnForClosePrice As Integer
    Dim ColumnForVolume As Integer
    Dim StartRowForPrint As Integer
    Dim ColumnForTickerPrint As Integer
    Dim ColumnForYearlyChange As Integer
    Dim ColumnForPercentChange As Integer
    Dim ColumnForTotalStockVolume As Integer
    
    Dim TickerName As String
    Dim OpenPrice As Variant
    Dim ClosePrice As Variant
    Dim TotalStockVolume As Variant
    
    Dim YearlyChange As Range
    
    Dim StartTime As Single
    StartTime = Timer
    
    
    For i = 1 To Sheets.Count
    
        Sheets(i).Activate
        Cells.Select
        Cells.FormatConditions.Delete
        Cells.EntireColumn.AutoFit
        
    
        ' Initialize variables
        YearlyChangeStartRow = 2
        StartRowForParsing = 2
        StartRowForPrint = 2
        StartRowForSummary = 2
        ColumnForTicker = 1
        ColumnForOpenPrice = 3
        ColumnForClosePrice = 6
        ColumnForVolume = 7
        
        
        ColumnForTickerPrint = 9
        ColumnForYearlyChange = 10
        ColumnForPercentChange = 11
        ColumnForTotalStockVolume = 12
        
        ColumnForSummaryHeadline = ColumnForTotalStockVolume + 3
        
        GreatestIncreaseTicker = ""
        GreatestDecreaseTicker = ""
        GreatestVolumeTicker = ""
        GreatestIncreaseValue = 0
        GreatestDecreaseValue = 0
        GreatestTotalVolume = 0
        
        
        Cells(StartRowForPrint - 1, ColumnForTickerPrint).Value = "Ticker"
        Cells(StartRowForPrint - 1, ColumnForYearlyChange).Value = "Yearly Change"
        Cells(StartRowForPrint - 1, ColumnForPercentChange).Value = "Percent Change"
        Cells(StartRowForPrint - 1, ColumnForTotalStockVolume).Value = "Total Stock Volume"
        
        Cells(StartRowForSummary - 1, ColumnForSummaryHeadline + 1).Value = "Ticker"
        Cells(StartRowForSummary - 1, ColumnForSummaryHeadline + 2).Value = "Value"
        Cells(StartRowForSummary, ColumnForSummaryHeadline).Value = "Greatest % Increase"
        Cells(StartRowForSummary + 1, ColumnForSummaryHeadline).Value = "Greatest % Decrease"
        Cells(StartRowForSummary + 2, ColumnForSummaryHeadline).Value = "Greatest Total Volume"
        
        
        TotalStockVolume = 0
        
        Let TickerSymbol = Cells(StartRowForParsing, ColumnForTicker).Value
    
        While (TickerSymbol <> "") 'Break when there are no more stocks in table
        
            TickerName = Cells(StartRowForParsing, ColumnForTicker).Value
            OpenPrice = Cells(StartRowForParsing, ColumnForOpenPrice).Value
            TotalStockVolume = TotalStockVolume + Cells(StartRowForParsing, ColumnForVolume).Value
            
            StartRowForParsing = StartRowForParsing + 1
            
            While (TickerSymbol = Cells(StartRowForParsing, ColumnForTicker).Value) 'Break when a new ticker name is found
                TotalStockVolume = TotalStockVolume + Cells(StartRowForParsing, ColumnForVolume).Value
                StartRowForParsing = StartRowForParsing + 1
            Wend
            ClosePrice = Cells(StartRowForParsing - 1, ColumnForClosePrice).Value
            
            ' Print Stock Summary
        
            Cells(StartRowForPrint, ColumnForTickerPrint).Value = TickerName
            Cells(StartRowForPrint, ColumnForYearlyChange).Value = ClosePrice - OpenPrice
            
            Let PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)
            
            Cells(StartRowForPrint, ColumnForPercentChange).Value = PercentChange
            Cells(StartRowForPrint, ColumnForPercentChange).Style = "Percent"
            Cells(StartRowForPrint, ColumnForPercentChange).NumberFormat = "0.00%"
            
            Cells(StartRowForPrint, ColumnForTotalStockVolume).Value = TotalStockVolume
            
            If StartRowForPrint = 2 Then
                GreatestIncreaseTicker = TickerName
                GreatestDecreaseTicker = TickerName
                GreatestVolumeTicker = TickerName
                GreatestIncreaseValue = PercentChange
                GreatestDecreaseValue = PercentChange
                GreatestTotalVolume = TotalStockVolume
            End If
            
            If PercentChange > GreatestIncreaseValue Then
                GreatestIncreaseValue = PercentChange
                GreatestIncreaseTicker = TickerName
            ElseIf PercentChange < GreatestDecreaseValue Then
                GreatestDecreaseValue = PercentChange
                GreatestDecreaseTicker = TickerName
            End If
            
            If TotalStockVolume > GreatestTotalVolume Then
                GreatestTotalVolume = TotalStockVolume
                GreatestVolumeTicker = TickerName
            End If
            
            StartRowForPrint = StartRowForPrint + 1
            TotalStockVolume = 0
            ' End Print Stock Summary
                    
            TickerSymbol = Cells(StartRowForParsing, ColumnForTicker).Value
        Wend
        
        'Conditional Formatting
        Set YearlyChange = Range(Cells(YearlyChangeStartRow, ColumnForYearlyChange), Cells(StartRowForPrint - 1, ColumnForYearlyChange))
        
        YearlyChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        YearlyChange.FormatConditions(1).Interior.Color = vbGreen
        
        YearlyChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        YearlyChange.FormatConditions(2).Interior.Color = vbRed
        
        YearlyChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
            Formula1:="=0"
        YearlyChange.FormatConditions(3).Interior.Color = vbBlue
        
        Cells(StartRowForSummary, ColumnForSummaryHeadline + 1).Value = GreatestIncreaseTicker
        Cells(StartRowForSummary + 1, ColumnForSummaryHeadline + 1).Value = GreatestDecreaseTicker
        Cells(StartRowForSummary + 2, ColumnForSummaryHeadline + 1).Value = GreatestVolumeTicker
        Cells(StartRowForSummary, ColumnForSummaryHeadline + 2).Value = GreatestIncreaseValue
        Cells(StartRowForSummary, ColumnForSummaryHeadline + 2).Style = "Percent"
        Cells(StartRowForSummary, ColumnForSummaryHeadline + 2).NumberFormat = "0.00%"
        Cells(StartRowForSummary + 1, ColumnForSummaryHeadline + 2).Value = GreatestDecreaseValue
        Cells(StartRowForSummary + 1, ColumnForSummaryHeadline + 2).Style = "Percent"
        Cells(StartRowForSummary + 1, ColumnForSummaryHeadline + 2).NumberFormat = "0.00%"
        Cells(StartRowForSummary + 2, ColumnForSummaryHeadline + 2).Value = GreatestTotalVolume
        
        Cells(1, 1).Select
    
    Next i
    
    'Time taken by code in seconds
    MsgBox Timer - StartTime
    
End Sub




