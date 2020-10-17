Attribute VB_Name = "working_stock_table"
Option Explicit

Sub stock_table()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
            
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percentage Change"
    Cells(1, 12) = "Total Stock Volume"
  
    Dim name As String
    Dim opn, clse, percent_change, yearly_change As Double
    Dim r, c, volumn, lrow, table_row  As LongLong
    
    'find last row value
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'what is the opening stock value
    opn = Cells(2, 3).Value
    
    'table row number
    table_row = 2
    
        For r = 2 To lrow
    
            If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
                name = Cells(r, 1).Value
                Cells(table_row, 9) = name
                
                'Find yearly change for close to open
                clse = Cells(r, 6).Value
                yearly_change = clse - opn
                Cells(table_row, 10).Value = yearly_change
                
                'conditional green red yearly_change
                    If Cells(table_row, 10).Value >= 0 Then
                        Cells(table_row, 10).Interior.ColorIndex = 4
                        Else: Cells(table_row, 10).Interior.ColorIndex = 3
                    End If
                
                'Find % in change, if opn zero, then zero %
                    If opn = 0 Then
                        Cells(table_row, 11) = 0
                    Else: percent_change = yearly_change / opn
                        Cells(table_row, 11) = percent_change
                        Cells(table_row, 11).NumberFormat = "0.000%"
                        
                    End If
                
                'total volumn
                volumn = volumn + Cells(r, 7).Value
                Cells(table_row, 12) = volumn
                
                'reset stock volumn
                volumn = 0
                
                'add row to table
                table_row = table_row + 1
                
                'new open value for next stock
                opn = Cells(r + 1, 3).Value
            
                'rolling volumn total until unequal ticker
                Else: volumn = volumn + ActiveSheet.Cells(r, 7).Value
        
            End If
        
        
        Next r


    'Finding Max, Min and Total
    
        Cells(2, 14) = "Greatest % Increase"
        Cells(3, 14) = "Greatest % Decrease"
        Cells(4, 14) = "Greatest Total Volumn"
        Cells(1, 15) = "Ticker"
        Cells(1, 16) = "Value"
        
        Dim percentage As Double
        Dim total_volumn
        
        'find last row new table
        lrow = Cells(Rows.Count, 9).End(xlUp).Row
        
        'set firt % & define as place holder
        percentage = Cells(2, 11).Value
        
        For c = 2 To lrow - 1
            If percentage < Cells(c + 1, 11).Value Then
                percentage = Cells(c + 1, 11).Value
                'ticker name
                Cells(2, 15).Value = Cells(c + 1, 9).Value
                ' percentage value increase
                Cells(2, 16).Value = Cells(c + 1, 11).Value
                Cells(2, 16).NumberFormat = "0.000%"
            End If
        Next c
        
        'find % decrease
        'rest value
        percentage = Cells(2, 11).Value
        For c = 2 To lrow - 1
            If percentage > Cells(c + 1, 11).Value Then
                percentage = Cells(c + 1, 11).Value
                'ticker name
                Cells(3, 15).Value = Cells(c + 1, 9).Value
                ' percentage value decrease
                Cells(3, 16).Value = Cells(c + 1, 11).Value
                Cells(3, 16).NumberFormat = "0.000%"
            End If
        Next c
            
        'Find greatest volum
        percentage = Cells(2, 12).Value
        For c = 2 To lrow - 1
            If percentage < Cells(c + 1, 12).Value Then
                percentage = Cells(c + 1, 12).Value
                'ticker name
                Cells(4, 15).Value = Cells(c + 1, 9).Value
                ' percentage value decrease
                Cells(4, 16).Value = Cells(c + 1, 12).Value
            End If
        Next c
        
    Next ws
    
End Sub
