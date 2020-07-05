Sub alpha_test()
 
'loop through worksheets
'?????
    
'set variables for ticker
Dim ticker As String

Dim year_open As Double
Dim year_close As Double

Dim year_change As Double
year_change = 0

Dim summary_row As Integer
summary_row = 2

Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Dim percent_change As Double

'set first ticker and year_open value
For i = 2 To last_row
    

    If i = 2 Then
        year_open = Cells(2, 3).Value
        ticker = Cells(2, 1).Value
        'Range("K2").Value = year_open
        
    End If
    
    'set the rest of the tickers, year_open, and year_close
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'find year close number
    year_close = Cells(i, 6).Value
    'print year_close to check accuracy
    'Range("L" & summary_row).Value = year_close
    
    'set the ticker
    ticker = Cells(i, 1).Value
    
    'print the ticker in the summary table
    Range("J" & summary_row) = ticker
    
   'find year change
    year_change = year_close - year_open
    'print year change
    Range("K" & summary_row) = year_change
    
    'Reduce possibility of getting math error
    If (year_open > 0) And (year_change <> 0) Then
    'find percent change
    percent_change = ((year_change) / (year_open)) * 100
    End If
    
    'print percent_change to summary table
    Range("L" & summary_row).Value = percent_change
   
   'find year open for next cell
    year_open = Cells(i + 1, 3).Value
    'print year_close to check accuracy
    'Range("K" & (summary_row + 1)) = year_open
    
    'advance row in summary table
    summary_row = summary_row + 1
    
    End If
    
    Next i
  
  
  'total volume loop setup
  
    summary_row = 2
    
        Dim total_volume As Variant
    
    For K = 2 To last_row
    
    
    If Cells(K, 1).Value <> Cells(K + 1, 1).Value Then
    
    'add last volume tick before change
    total_volume = total_volume + Cells(K, 7).Value
      
    'print total volume to sheet
    Range("M" & summary_row).Value = total_volume
    
    'add one to summary_row
    summary_row = summary_row + 1
    
   'reset total volume
    total_volume = 0
    
    Else
    
    'otherwise just keep adding!
        total_volume = total_volume + Cells(K, 7).Value

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
    
    If (Cells(m + 1, 12).Value > Cells(m, 12).Value) And (Cells(m + 1, 12).Value > best_increase) Then
    
    best_increase = Cells(m + 1, 12).Value
    best_ticker = Cells(m + 1, 10).Value
    
    End If
    
    If (Cells(m + 1, 12).Value < Cells(m, 12).Value) And (Cells(m + 1, 12).Value < worst_decrease) Then
    
    worst_decrease = Cells(m + 1, 12).Value
    worst_ticker = Cells(m + 1, 10).Value
    
    End If
    
    If (Cells(m + 1, 13).Value > Cells(m, 13).Value) And (Cells(m + 1, 13).Value > greatest_volume) Then

    
    greatest_volume = Cells(m + 1, 13).Value
    greatest_ticker = Cells(m + 1, 10).Value
    
    End If
    
    Next m
    
    'print summaries
    Range("O" & summary_table).Value = "Greatest % Increase"
    Range("P" & summary_table).Value = best_ticker
    Range("Q" & summary_table).Value = best_increase
    
    'move down one row on summary table
    summary_table = summary_table + 1
    
    'print negative summary
    Range("O" & summary_table).Value = "Greatest % Decrease"
    Range("P" & summary_table).Value = worst_ticker
    Range("Q" & summary_table).Value = worst_decrease
    
    summary_table = summary_table + 1
    
    'print volume summary
    Range("O" & summary_table).Value = "Greatest Total Volume"
    Range("P" & summary_table).Value = greatest_ticker
    Range("Q" & summary_table).Value = greatest_volume
    
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
    
End Sub



