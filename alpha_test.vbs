Sub alpha_test()
 
'loop through worksheets
'?????

Dim ws As Worksheet

For Each ws In Worksheets

'set variables for ticker
Dim ticker As String

Dim year_open As Double
Dim year_close As Double

Dim year_change As Double
year_change = 0

Dim summary_row As Integer
summary_row = 2

Dim last_row As Long
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim percent_change As Double

'set first ticker and year_open value
For i = 2 To last_row
    

    If i = 2 Then
        year_open = ws.Cells(2, 3).Value
        ticker = ws.Cells(2, 1).Value
        'Range("K2").Value = year_open
        
    End If
    
    'set the rest of the tickers, year_open, and year_close
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'find year close number
    year_close = ws.Cells(i, 6).Value
    'print year_close to check accuracy
    'Range("L" & summary_row).Value = year_close
    
    'set the ticker
    ticker = ws.Cells(i, 1).Value
    
    'print the ticker in the summary table
    ws.Range("J" & summary_row) = ticker
    
   'find year change
    year_change = year_close - year_open
    'print year change
    ws.Range("K" & summary_row) = year_change
    
    'Reduce possibility of getting math error
    If (year_open > 0) And (year_change <> 0) Then
    'find percent change
    percent_change = ((year_change) / (year_open)) * 100
    End If
    
    'print percent_change to summary table
    ws.Range("L" & summary_row).Value = percent_change
   
   'find year open for next cell
    year_open = ws.Cells(i + 1, 3).Value
    'print year_close to check accuracy
    'Range("K" & (summary_row + 1)) = year_open
    
    'advance row in summary table
    summary_row = summary_row + 1
    
    End If
    
    Next i
  
  
  'total volume loop setup
  
    summary_row = 2
    
        Dim total_volume As Variants
    
    For K = 2 To last_row
    
    
    If ws.Cells(K, 1).Value <> ws.Cells(K + 1, 1).Value Then
    
    'add last volume tick before change
    total_volume = total_volume + ws.Cells(K, 7).Value
      
    'print total volume to sheet
    ws.Range("M" & summary_row).Value = total_volume
    
    'add one to summary_row
    summary_row = summary_row + 1
    
   'reset total volume
    total_volume = 0
    
    Else
    
    'otherwise just keep adding!
        total_volume = total_volume + ws.Cells(K, 7).Value

    End If
    
    Next K
    
    'set up loop to find challenge answers
    For m = 2 To last_row
    
    'set variable for greatest % increase
    Dim best_increase As Double
    Dim best_ticker As String
    Dim summary_table As Variant
    Dim worst_ticker As String
    Dim worst_decrease As Double
    Dim greatest_volume As Variant
    Dim greatest_ticker As String
    
    
    summary_table = 2
    
    If (ws.Cells(m + 1, 12).Value > ws.Cells(m, 12).Value) And (ws.Cells(m + 1, 12).Value > best_increase) Then
    
    best_increase = ws.Cells(m + 1, 12).Value
    best_ticker = ws.Cells(m + 1, 10).Value
    
    End If
    
    If (ws.Cells(m + 1, 12).Value < ws.Cells(m, 12).Value) And (ws.Cells(m + 1, 12).Value < worst_decrease) Then
    
    worst_decrease = ws.Cells(m + 1, 12).Value
    worst_ticker = ws.Cells(m + 1, 10).Value
    
    End If
    
    If (ws.Cells(m + 1, 13).Value > ws.Cells(m, 13).Value) And (ws.Cells(m + 1, 13).Value > greatest_volume) Then

    
    greatest_volume = ws.Cells(m + 1, 13).Value
    greatest_ticker = ws.Cells(m + 1, 10).Value
    
    End If
    
    Next m
    
    'print summaries
    ws.Range("O" & summary_table).Value = "Greatest % Increase"
    ws.Range("P" & summary_table).Value = best_ticker
    ws.Range("Q" & summary_table).Value = best_increase
    
    'move down one row on summary table
    summary_table = summary_table + 1
    
    'print negative summary
    ws.Range("O" & summary_table).Value = "Greatest % Decrease"
    ws.Range("P" & summary_table).Value = worst_ticker
    ws.Range("Q" & summary_table).Value = worst_decrease
    
    summary_table = summary_table + 1
    
    'print volume summary
    ws.Range("O" & summary_table).Value = "Greatest Total Volume"
    ws.Range("P" & summary_table).Value = greatest_ticker
    ws.Range("Q" & summary_table).Value = greatest_volume
    
     'Label Columns
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    
    'conditional formatting (must be in cell A1 :/ )
    ActiveCell.Offset(0, 10).Columns("A:A").EntireColumn.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
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
    ActiveCell.Offset(1, 0).Range("A1").Activate
    
    
Next ws
End Sub





