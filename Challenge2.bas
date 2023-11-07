Attribute VB_Name = "Challenge2"
Option Explicit

Sub ProcessSheets()
    Dim i As Integer
    Application.ScreenUpdating = False
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        
        Sheets(ThisWorkbook.Worksheets(i).Name).Select
        'Clear prior
        Columns("I:Q").Delete
        
        'Add headers for output
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change ($)"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'Actual work
        DoTheWork (ThisWorkbook.Worksheets(i).Name)
        
        'Some formatting to make it look prettier
        Columns("K:K").Select
        Selection.NumberFormat = "0.00%"
        
        Range("Q2:Q3").Select
        Selection.NumberFormat = "0.00%"
        
        Columns("I:Q").EntireColumn.AutoFit
        
        'add conditional formatting, red for < 0 and green for >= 0
        Range("J2:K2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.End(xlUp).Select
        
        'freeze header row
        Rows("1:1").Select
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
        
        Range("A2").Select
    Next i
    
    Sheets(ThisWorkbook.Worksheets(1).Name).Select
    
    Application.ScreenUpdating = True
    
    MsgBox "Done processing " & ThisWorkbook.Worksheets.Count & " sheets"
    
End Sub

Sub DoTheWork(inSheetName As String)
    'Variables for data table
    Dim sTicker As String
    Dim dOpen As Double
    Dim dClose As Double
    Dim dVolume As Double
    
    'variables for output
    Dim dYearlyChange As Double
    Dim dPercentChange As Double
    
    'variables to keep track of rows
    Dim dDataRow As Double
    Dim iOutputRow As Integer
    Dim dLastDataRow As Double
    
    'variable to check for new ticker
    Dim sNextTicker As String
    
    'Move to the right worksheet
    Sheets(inSheetName).Select
    
    'initialize variables
    dDataRow = 2
    iOutputRow = 2
    sTicker = ""
    
    'Find the last data row
    dLastDataRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("A2").Select
    
    'Loop through each ticker row
    Do While dDataRow <= dLastDataRow
        If Len(sTicker) = 0 Then    'This is a new ticker, get ticker name and open value
            sTicker = Cells(dDataRow, 1).Value
            dOpen = Cells(dDataRow, 3).Value
        End If
        
        'this will just keep overwriting the close value until we reach the last found
        dClose = Cells(dDataRow, 6).Value
        
        'add to existing to sum all totals
        dVolume = dVolume + Cells(dDataRow, 7).Value
        
        'look ahead to see if there's a different ticker value
        
        sNextTicker = Cells(dDataRow + 1, 1).Value
        
        If sTicker <> sNextTicker Then  'Time to write to the output
            dYearlyChange = dClose - dOpen
            dPercentChange = (dClose - dOpen) / dOpen
            
            Cells(iOutputRow, Columns("I").Column).Value = sTicker
            Cells(iOutputRow, Columns("J").Column).Value = dYearlyChange
            Cells(iOutputRow, Columns("K").Column).Value = dPercentChange
            Cells(iOutputRow, Columns("L").Column).Value = dVolume
            
            'Check for greatests table
            'greatest % increase
            If dPercentChange > Range("Q2").Value Then
                Range("P2").Value = sTicker
                Range("Q2").Value = dPercentChange
            End If
            'greatest % decrease
            If dPercentChange < Range("Q3").Value Then
                Range("P3").Value = sTicker
                Range("Q3").Value = dPercentChange
            End If
            'greatest total volume
            If dVolume > Range("Q4").Value Then
                Range("P4").Value = sTicker
                Range("Q4").Value = dVolume
            End If
            
            'Reset variables
            sTicker = ""
            dOpen = 0
            dClose = 0
            dVolume = 0
            dYearlyChange = 0
            dPercentChange = 0
            sNextTicker = ""
            
            'increment output row
            iOutputRow = iOutputRow + 1
        End If
        
        'increment row
        dDataRow = dDataRow + 1
    Loop
End Sub





