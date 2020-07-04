Sub alpha_test()

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

For i = 2 To last_row
    

    If i = 2 Then
        year_open = Cells(2, 3).Value
        ticker = Cells(2, 1).Value
        'Range("K2").Value = year_open
        
    End If
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'find year close number
    year_close = Cells(i, 6).Value
    'Range("L" & summary_row).Value = year_close
    
    
    'set the ticker
    ticker = Cells(i, 1).Value
    
      'put the ticker in the summary table
    Range("J" & summary_row) = ticker
    
   'find year change
    year_change = year_close - year_open
    Range("K" & summary_row) = year_change
    
    'find percent change
    percent_change = ((year_change) / (year_open)) * 100
   Range("L" & summary_row).Value = percent_change
   
   'find year open for next cell
    year_open = Cells(i + 1, 3).Value
    'Range("K" & (summary_row + 1)) = year_open
    
    'advance row in summary table
    summary_row = summary_row + 1
    
    End If
    
    'add final ticker and final year_close
    
  
    
    
    Next i
  
    summary_row = 2
    
        Dim total_volume As Variant
    
    For k = 2 To last_row
    
    
    If Cells(k, 1).Value <> Cells(k + 1, 1).Value Then
    
    'add last volume tick before change
    total_volume = total_volume + Cells(k, 7).Value
      
    'print total volume to sheet
    Range("M" & summary_row).Value = total_volume
    
    'add one to summary_row
    summary_row = summary_row + 1
    
   'reset total volume
    total_volume = 0
    
    Else
    
    'otherwise just keep adding!
        total_volume = total_volume + Cells(k, 7).Value
        
        'getting overflow error here ^^ but first volume is
        'printed correctly for ticker A... Set dim total_volume as variant
        'and that seemed to solve the issue!
    
    End If
    
    Next k
    
    


End Sub
