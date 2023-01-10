Attribute VB_Name = "Module1"
Sub run_all()

Dim ws As Worksheet

For Each ws In Worksheets

' Activate current worksheet
ws.Activate

' Call on programs
clear_sheet
summary_table
conditionals_formatting
summary_table_2

Next ws

End Sub

Sub summary_table()

Dim last_row As Double
Dim row_counter As Double
Dim ticker As String
Dim yearly_change As Double
Dim open_price As Double
Dim close_price As Double
Dim percentage_change As Double
Dim total_stock_volume As Double
Dim open_price_captured As Boolean
    
' Set number of rows of data in sheet
last_row = Cells(Rows.Count, "A").End(xlUp).Row
'MsgBox (last_row)

' Establish new columns and rows
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"

' Set initial values
row_counter = 2
yearly_change = 0
open_price = 0
close_price = 0
percentage_change = 0
total_stock_volume = 0

    ' Loop through all rows
    For i = 2 To last_row

    ticker = Cells(i, 1).Value
        
        ' Capture inital opening price
        If open_price_captured = False Then
            
            open_price = Cells(i, 3).Value

            open_price_captured = True
            
        End If

        ' Calculate yearly change, percent change and total stock volume
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Copy unique tickers into ticker column
            Cells(row_counter, 9).Value = ticker
            
            ' Capture closing price and calculate Yearly Change
            close_price = Cells(i, 6).Value
            yearly_change = close_price - open_price
            Cells(row_counter, 10).Value = yearly_change
            
            ' Calculate Percentage Change
            percentage_change = (yearly_change / open_price)
            Cells(row_counter, 11).Value = percentage_change
            
            ' Calculate sum of individual stocks
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            Cells(row_counter, 12).Value = total_stock_volume
            
            row_counter = row_counter + 1
            total_stock_volume = 0
            open_price_captured = False
            
            ' Continue sum of stock total
            Else
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
        End If
        
    Next i
                
End Sub

Sub conditionals_formatting()

Dim last_row2 As Integer

' Counter number of rows for the first summary table
last_row2 = Cells(Rows.Count, 9).End(xlUp).Row

    ' Loop through first summary table
    For i = 2 To last_row2
    
        ' Format cell colors
        If Cells(i, 10).Value > 0 Then
            
            ' Format positive values as green
            Cells(i, 10).Interior.ColorIndex = 4
            
        ElseIf Cells(i, 10).Value < 0 Then
            
            ' Format negative values as red
            Cells(i, 10).Interior.ColorIndex = 3
        
        End If
            
    Next i
    
' Format Percentage Change and Total Stock Volume columns
Range("J1:J" & last_row2).NumberFormat = "$#,##0.00"
Range("K1:K" & last_row2).NumberFormat = "0.00%"
Range("L1:L" & last_row2).NumberFormat = "#,##0"

' Autofit columns to data
ActiveSheet.UsedRange.EntireColumn.AutoFit
        
End Sub
        
Sub summary_table_2()
        
Dim last_row2 As Integer
Dim ticker2 As String
Dim ticker3 As String
Dim ticker4 As String
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_total_volume As Double

' Set column and row labels
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest Total Volume"

' Counter number of rows for the first summary table
last_row2 = Cells(Rows.Count, 9).End(xlUp).Row

' Set initial values for comparison
greatest_increase = -1000
greatest_decrease = 1000

    ' Loop through first summary table
    For i = 2 To last_row2
        
        ' Find greatest % increase and paste data
        If greatest_increase < Cells(i, 11).Value Then
            
            greatest_increase = Cells(i, 11).Value
            ticker2 = Cells(i, 9).Value
    
        End If
    
        If Cells(i, 11) < greatest_decrease Then
            
            greatest_decrease = Cells(i, 11).Value
            ticker3 = Cells(i, 9).Value
    
        End If
        
        If greatest_total_volume < Cells(i, 12).Value Then
            
            greatest_total_volume = Cells(i, 12).Value
            ticker4 = Cells(i, 9).Value
    
        End If
    
    Next i

' Print all values
Range("P2").Value = ticker2
Range("P3").Value = ticker3
Range("P4").Value = ticker4

Range("Q2").Value = greatest_increase
Range("Q3").Value = greatest_decrease
Range("Q4").Value = greatest_total_volume

' Format cells, percentage, round and comma
Range("Q2:Q3").NumberFormat = "0.00%"
Range("Q4").NumberFormat = "#,##0"

' Autofit columns to data
ActiveSheet.UsedRange.EntireColumn.AutoFit

End Sub

Sub clear_sheet()

Range("H:Q").Clear

End Sub

Sub clear_all()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Activate

Range("H:Q").Clear

Next ws

End Sub


