Attribute VB_Name = "AlphabeticalStockLAS"
Sub sheet_selected()
'--> main routine that navigates on every sheet of the book
'--> declare variables
Dim sheet_counter As Integer
'--> to know number of sheets in the book
sheet_counter = ThisWorkbook.Sheets.Count
    For i = 1 To sheet_counter '--> do for each sheet
        ThisWorkbook.Sheets(i).Select
        start_proccess (i) '--> calculates and prints values grouped by ticker
        Greatest_Values    '--> routine for calculating the greater values
    Next i
    MsgBox ("END")
    ThisWorkbook.Sheets(1).Select
End Sub
Sub start_proccess(sheet_ As Integer)
'--> subroutine that calculates and prints values grouped by ticker
'--> declare variables
Dim Total_rows As LongLong
Dim total_stock_volume As LongLong
Dim Yearly_change As Double
Dim ticker_symbol As String
Dim percent_change As Double
Dim changecolorcell As String
Dim total_open_price As Double
Dim total_close_price As Double
Dim ticker_count As LongLong
Dim area_clear As String

'-->inicialize variables
ticker_count = 1
total_open_price = 0#
total_close_price = 0#
total_stock_volume = 0

'--> selected sheet to make the calculations
ThisWorkbook.Sheets(sheet_).Select
'--> identify number of total rows consecutive with data
Total_rows = Cells(Rows.Count, 1).End(xlUp).Row

'--> clean up area of data print, columns i to o
area_clear = "I1:R" + CStr(Total_rows)
Sheets(sheet_).Range(area_clear).Clear
Sheets(sheet_).Range("I1").Select

'--> headers / titles
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To Total_rows '--> Total_rows
        If i = 2 Then  '-->save dato of row 2 to camparate with the next rows
            ticker_symbol = Cells(i, 1).Value
            total_open_price = Cells(i, 3).Value
            total_close_price = Cells(i, 6).Value
            total_stock_volume = Cells(i, 7).Value
            ticker_count = ticker_count + 1
        Else
            If ticker_symbol = Cells(i, 1).Value Then
             '--> if so equal then just acumulate values for the same ticker
               total_open_price = total_open_price + Cells(i, 3).Value
               total_close_price = total_close_price + Cells(i, 6).Value
               total_stock_volume = total_stock_volume + Cells(i, 7).Value
            Else
                '-->diferent ticker then print last ticker
                Cells(ticker_count, 9).Value = ticker_symbol
                Cells(ticker_count, 10).Value = (total_close_price - total_open_price)
                If Cells(ticker_count, 10).Value > 0 Then
                    Cells(ticker_count, 10).Interior.ColorIndex = 4 '--> cells color equal to green
                Else
                    Cells(ticker_count, 10).Interior.ColorIndex = 3 '--> cells color equal to red
                End If
                If total_open_price <> 0 Then
                    Cells(ticker_count, 11).Value = ((total_close_price - total_open_price) / total_open_price)
                Else
                    Cells(ticker_count, 11).Value = 0
                End If
                Cells(ticker_count, 12).Value = total_stock_volume
                'inicialize variables for next sticker
                total_open_price = 0#
                total_close_price = 0#
                total_stock_volume = 0
                '-->acumulate new ticker
                ticker_symbol = Cells(i, 1).Value
                total_open_price = Cells(i, 3).Value
                total_close_price = Cells(i, 6).Value
                total_stock_volume = Cells(i, 7).Value
                ticker_count = ticker_count + 1
            End If
        End If
Next i
'--> print last ticker values
Cells(ticker_count, 9).Value = ticker_symbol
Cells(ticker_count, 10).Value = (total_close_price - total_open_price)
If Cells(ticker_count, 10).Value > 0 Then
    Cells(ticker_count, 10).Interior.ColorIndex = 4 '--> cells color equal to green
Else
    Cells(ticker_count, 10).Interior.ColorIndex = 3 '--> cells color equal to red
End If
If total_open_price <> 0 Then
    Cells(ticker_count, 11).Value = ((total_close_price - total_open_price) / total_open_price)
Else
    Cells(ticker_count, 11).Value = 0
End If
Cells(ticker_count, 12).Value = total_stock_volume
Range("K2:K" & CStr(Cells(Rows.Count, 11).End(xlUp).Row)).NumberFormat = "0.00%" '--> express as a percentage column K

End Sub
Sub Greatest_Values()
'--> routine for calculating the greater values
'--> declare variables for greatest values
Dim data_range As Excel.Range
Dim max_Range As Excel.Range
Dim min_Range As Excel.Range
Dim max_stock As Excel.Range
Dim Greatest_increase As String
Dim Greatest_decrease As String
Dim Greatest_total_volume As LongLong
Dim delete_area As String
Dim c As Integer '--> position of the greatest
'--> initialize variables
c = 1
'--> Title Greatest values
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest total volume"

Application.Range("P1").Select
Set data_range = Range("K2:K" & CStr(Cells(Rows.Count, 10).End(xlUp).Row))
For Each max_Range In data_range
    If max_Range.Value = Application.WorksheetFunction.Max(data_range) Then
        c = c + 1
        'sAddress = max_Range.Address
        'vPrevValue = max_Range.Value
        cell_Ticker = "P" & c ' to print ticker value
        cell_value = "Q" & c  ' to print greatest value
        cell_Title = "O" & c  ' to print greatest Title
        Range(cell_Title).Value = "Greatest % increase"
        Range(cell_Ticker).Value = Range("I" & Range(max_Range.Address).Row).Value
        Range(cell_value).Value = Range(max_Range.Address).Value
        Range(cell_value).NumberFormat = "0.00%" '--> express as a percentage
    End If
Next max_Range
For Each min_Range In data_range
    If min_Range.Value = Application.WorksheetFunction.Min(data_range) Then
        c = c + 1
        'sAddress = min_Range.Address
        'vPrevValue = min_Range.Value
        cell_Ticker = "P" & c ' to print ticker value
        cell_value = "Q" & c  ' to print greatest value
        cell_Title = "O" & c  ' to print greatest Title
        Range(cell_Title).Value = "Greatest % decrease"
        Range(cell_Ticker).Value = Range("I" & Range(min_Range.Address).Row).Value
        Range(cell_value).Value = Range(min_Range.Address).Value
        Range(cell_value).NumberFormat = "0.00%" '--> express as a percentage
    End If
Next min_Range
Set data_range = Range("L2:L" & CStr(Cells(Rows.Count, 9).End(xlUp).Row))
For Each max_stock In data_range
    If max_stock.Value = Application.WorksheetFunction.Max(data_range) Then
        c = c + 1
        'sAddress = max_stock.Address
        'vPrevValue = max_stock.Value
        cell_Title = "O" & c  ' to print greatest Title
        cell_Ticker = "P" & c ' to print ticker value
        cell_value = "Q" & c  ' to print greatest value
        Range(cell_Title).Value = "Greatest total volume"
        Range(cell_Ticker).Value = Range("I" & Range(max_stock.Address).Row).Value
        Range(cell_value).Value = Range(max_stock.Address).Value
    End If
Next max_stock
End Sub

