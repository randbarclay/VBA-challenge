Attribute VB_Name = "Module1"
Sub vba_challenge()

'for loop to loop through all worksheets...

For Each ws In Worksheets


'declaring and defining our variables...

Dim lastrow As Long
Dim total_volume As Double
Dim annual_change As Double
Dim annual_percent_change As Double
Dim stock_name As String
Dim summary_table_row As Integer
Dim first_price As Double
Dim last_price As Double

summary_table_row = 2
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row


'formatting the headers...

ws.Range("I1").Value = "Ticker Symbol"
ws.Range("J1").Value = "Annual Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'for loop to create the table of values...

For i = 2 To lastrow

If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

    first_price = ws.Cells(i, 6).Value
    
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        stock_name = ws.Cells(i, 1).Value
        total_volume = total_volume + ws.Cells(i, 7).Value
        last_price = ws.Cells(i, 6).Value
        ws.Range("I" & summary_table_row).Value = stock_name
        ws.Range("L" & summary_table_row).Value = total_volume
        ws.Range("J" & summary_table_row).Value = first_price - last_price
        
        'nest an if inside this for loop to account for any possible 0 stock values...
        
        If last_price = 0 Then
        
        annual_percent_change = 0
        
        Else
        
        annual_percent_change = ws.Range("J" & summary_table_row).Value / last_price
        
        End If
        
        ws.Range("K" & summary_table_row).Value = annual_percent_change
        ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
        summary_table_row = summary_table_row + 1
        total_volume = 0
        
        Else
    
        total_volume = total_volume + ws.Cells(i, 7).Value
    
    End If

Next i

'for loop to properly format yearly change per the instructions...

Dim lastrow2 As Long
lastrow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row

For i = 2 To lastrow2

If ws.Cells(i, 10).Value < 0 Then
    
    ws.Cells(i, 10).Interior.ColorIndex = 3
    
    'the assignment says to only format postive/negative values...meaning we do not
    'want to format 0 values. this elseif will do nothing if the cell is zero
    
    ElseIf ws.Cells(i, 10).Value = 0 Then
    
    Else
    
    ws.Cells(i, 10).Interior.ColorIndex = 4

End If

Next i

'formatting the max/min table...

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker Symbol"
ws.Range("Q1").Value = "Value"

'obtaining max/min values...

ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K:K"))
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K:K"))
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L:L"))

'for loop to match the ticker symbol to the max/min figures...

For i = 2 To lastrow

If ws.Cells(i, 11).Value = ws.Range("Q2").Value Then
    ws.Range("P2").Value = ws.Cells(i, 9).Value
    
    End If
    
If ws.Cells(i, 11).Value = ws.Range("Q3").Value Then
    ws.Range("P3").Value = ws.Cells(i, 9).Value
    
    End If

If ws.Cells(i, 12).Value = ws.Range("Q4").Value Then
    ws.Range("P4").Value = ws.Cells(i, 9).Value
    
    End If
    
Next i

Next ws

End Sub


