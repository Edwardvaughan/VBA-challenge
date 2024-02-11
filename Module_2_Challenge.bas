Attribute VB_Name = "Module1"
Sub stocks()

' A program for giving a summary of stock data, written in VBA

    Dim y1 As Long
    Dim y2 As String
    Dim lastrow As Long
    Dim i As Long
    Dim j As Integer
    Dim ticker As String
    Dim previous_ticker As String
    Dim o As Double
    Dim previous_o As Double
    Dim c As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As LongLong
    Dim greatest_percent_increase_ticker As String
    Dim greatest_percent_increase As Double
    Dim greatest_percent_decrease_ticker As String
    Dim greatest_percent_decrease As Double
    Dim greatest_total_volume_ticker As String
    Dim greatest_total_volume As LongLong
    
    y1 = 2018
    
Step_1:

    y2 = CStr(y1)
    
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    previous_ticker = Worksheets(y2).Cells(2, 1).Value
    previous_o = Worksheets(y2).Cells(2, 3).Value
    total_stock_volume = Worksheets(y2).Cells(2, 7).Value
    
    Worksheets(y2).Cells(1, 9).Value = "Ticker"
    Worksheets(y2).Cells(1, 10).Value = "Yearly Change"
    Worksheets(y2).Cells(1, 11).Value = "Percent Change"
    Worksheets(y2).Cells(1, 12).Value = "Total Stock Volume"
    
    i = 3
    
    j = 2
    
    greatest_percent_increase = 0
    greatest_percent_decrease = 0
    greatest_total_volume = 0
    
' Get the current ticker and total stock volume
    
Step_2:

    ticker = Worksheets(y2).Cells(i, 1).Value

    If ticker = previous_ticker Then
    
        total_stock_volume = total_stock_volume + Worksheets(y2).Cells(i, 7).Value
    
' If we get to a new ticker then compute a summary based on the previous ticker

    Else
        
        o = Worksheets(y2).Cells(i, 3).Value
        c = Worksheets(y2).Cells(i - 1, 6).Value
        
        yearly_change = c - previous_o
        percent_change = yearly_change / previous_o * 100
        total_stock_volume = total_stock_volume - Worksheets(y2).Cells(i, 7).Value
        
' Display the summary next to the stock data
        
        Worksheets(y2).Cells(j, 9).Value = previous_ticker
        Worksheets(y2).Cells(j, 10).Value = yearly_change
        Worksheets(y2).Cells(j, 11).Value = percent_change
        Worksheets(y2).Cells(j, 12).Value = total_stock_volume
        
' Colour code the Yearly Change cells depending on whether positive
' or negative
        
        If yearly_change >= 0 Then
    
            Worksheets(y2).Cells(j, 10).Interior.Color = RGB(0, 255, 0)
            
        Else
        
            Worksheets(y2).Cells(j, 10).Interior.Color = RGB(255, 0, 0)
            
        End If
        
' Get summary statistics
        
        If percent_change > greatest_percent_increase Then
            
            greatest_percent_increase_ticker = ticker
            greatest_percent_increase = percent_change
            
        End If
        
        If percent_change < greatest_percent_decrease Then
        
            greatest_percent_decrease_ticker = ticker
            greatest_percent_decrease = percent_change
            
        End If
        
        If Worksheets(y2).Cells(i, 7).Value > greatest_total_volume Then
            greatest_total_volume_ticker = ticker
            greatest_total_volume = Worksheets(y2).Cells(i, 7).Value
            
        End If
        
        previous_ticker = ticker
        previous_o = o
        total_stock_volume = Worksheets(y2).Cells(i, 7).Value
        
        j = j + 1
        
    End If
    
' For the given years, until we get to the bottom of the spreadsheet,
' iterate through the rows

    If y1 < 2021 Then
    
        If i < lastrow + 1 Then
    
            i = i + 1
        
            GoTo Step_2
    
        Else
        
' Display summary statistics
        
            Worksheets(y2).Cells(1, 16).Value = "Ticker"
            Worksheets(y2).Cells(1, 17).Value = "Value"
            Worksheets(y2).Cells(2, 15).Value = "Greatest % Increase"
            Worksheets(y2).Cells(3, 15).Value = "Greatest % Decrease"
            Worksheets(y2).Cells(4, 15).Value = "Greatest Total Volume"
            
            Worksheets(y2).Cells(2, 16).Value = greatest_percent_increase_ticker
            Worksheets(y2).Cells(2, 17).Value = greatest_percent_increase
            Worksheets(y2).Cells(3, 16).Value = greatest_percent_decrease_ticker
            Worksheets(y2).Cells(3, 17).Value = greatest_percent_decrease
            Worksheets(y2).Cells(4, 16).Value = greatest_total_volume_ticker
            Worksheets(y2).Cells(4, 17).Value = greatest_total_volume
        
            y1 = y1 + 1
            
            If y1 = 2021 Then
            
                Exit Sub
                
            Else
            
                GoTo Step_1
        
            End If
        
        End If
        
    End If
                    
End Sub
