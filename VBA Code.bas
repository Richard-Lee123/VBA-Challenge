Attribute VB_Name = "Module1"
Sub alpha_test()

Dim ticker As String
Dim yearly_change As Double
Dim per_change As Double
Dim vol As Double
Dim i As Double
Dim lastrow As Double
Dim year_open As Double
Dim year_closed As Double
Dim new_row As Double


vol = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'testing lastrow
    'MsgBox ("lastrow:" & lastrow)

'Headers
Range("I1") = "Ticker"
Range("L1") = "Total Stock Volume"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

year_open = Cells(2, 3).Value
year_closed = 0
            
    'looping through data
    For i = 2 To lastrow
    

        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
       

            year_closed = Cells(i, 6).Value
               yearly_change = year_closed - year_open
               
            ticker = Cells(i, 1).Value
            
            vol = vol + Cells(i, 7).Value
                   
            
            per_change = yearly_change / year_open
            
            Range("I" & Summary_Table_Row).Value = ticker
            Range("L" & Summary_Table_Row).Value = vol
            Range("J" & Summary_Table_Row).Value = yearly_change
            Range("K" & Summary_Table_Row).Value = per_change
                        
                        
                If Cells(Summary_Table_Row, 10).Value >= 0 Then
                    Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            
                ElseIf Cells(Summary_Table_Row, 10).Value < 0 Then
                    Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            
            
                End If
            
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            ticker = 0
            vol = 0
            year_open = Cells(i + 1, 3).Value
            year_closed = 0
            
            If year_open = 0 Then
                For new_row = i To lastrow
                    If Cells(new_row, 3).Value <> 0 Then
                        year_open = Cells(new_row, 3).Value
                        Exit For
                        End If
                        Next new_row
                            End If
                            
        Else
        vol = vol + Cells(i, 7).Value
        
        End If
        
   Next i
   
   
   
Dim sheet As Worksheet
Dim max As Double
Dim min As Double
Dim gtVolume As Double

max = 0
min = 0
gtVolume = 0

Set sheet = ActiveSheet

Range("P1") = "Ticker"
Range("Q1") = "Value"
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"
    For i = 2 To lastrow
        With sheet.Cells(i, 11).Value
    
            If sheet.Cells(i, 11) > max Then
            max = sheet.Cells(i, 11)
            
            Range("Q2").Value = max
            
            ColumnName = Cells(i, 9).Value
                    Range("P2").Value = ColumnName
                    'MsgBox ColumnName
        
            
                
            End If

        End With
        
        
        
        With sheet.Cells(i, 11).Value
            If sheet.Cells(i, 11) < min Then
                min = sheet.Cells(i, 11)
        
                Range("Q3").Value = min
                
                ColumnName = Cells(i, 9).Value
                    Range("P3").Value = ColumnName
                    'MsgBox ColumnName
            End If
        End With
        
        
        With sheet.Cells(i, 12).Value
            If sheet.Cells(i, 12) > gtVolume Then
                gtVolume = sheet.Cells(i, 12)
        
                Range("Q4").Value = gtVolume
                
                ColumnName = Cells(i, 9).Value
                    Range("P4").Value = ColumnName
                    'MsgBox ColumnName
            End If
        End With
        
    Next i
    
Columns("K").NumberFormat = "0.00%"
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"

End Sub
