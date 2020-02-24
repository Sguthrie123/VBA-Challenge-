Attribute VB_Name = "Module1"
Sub WallStreet():

    
    For Each ws In Worksheets
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percentage Change "
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "Greatest % increase"
    ws.Cells(3, 15) = "Greatest % decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    
    Dim Yearly_change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    Dim Open_Price As Double
    Open_Price = 0
    Dim Close_Price As Double
    Close_Price = 0
    Dim Helper As Double
    Helper = 2
    Dim EOY_Counter As Integer
    EOY_Counter = 0
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_TV As Double
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_TV = 0
    
    
    
    For i = 1 To LastRow
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            ws.Cells(Helper, 9) = ws.Cells(i + 1, 1)
            Open_Price = ws.Cells(i + 1, 3)
            EOY_Counter = WorksheetFunction.CountIf(ws.Range("A:A"), ws.Cells(Helper, 9))
            Close_Price = ws.Cells(i + EOY_Counter, 6)
            ws.Cells(Helper, 10) = Close_Price - Open_Price
            If Open_Price = 0 Then
                ws.Cells(Helper, 11) = 0
            Else
                ws.Cells(Helper, 11) = ((Close_Price - Open_Price) / Open_Price) * 100
            End If
            
            Helper = Helper + 1
        End If
        
    Next i

LastRowColor = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim TSV_Counter As Integer
TSV_Counter = 2
Dim tik1 As String
Dim tik2 As String
Dim tik3 As String


    For j = 2 To LastRowColor
        If ws.Cells(j, 10) > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(j, 10) < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 0
        End If
        
        If ws.Cells(j, 11) > Greatest_Increase Then
            Greatest_Increase = ws.Cells(j, 11)
            tik1 = ws.Cells(j, 9)
        End If
        ws.Cells(2, 17) = Greatest_Increase
        ws.Cells(2, 16) = tik1
        
        If ws.Cells(j, 11) < Greatest_Decrease Then
            Greatest_Decrease = ws.Cells(j, 11)
            tik2 = ws.Cells(j, 9)
        End If
        ws.Cells(3, 17) = Greatest_Decrease
        ws.Cells(3, 16) = tik2
        
        If ws.Cells(j, 12) > Greatest_TV Then
            Greatest_TV = ws.Cells(j, 12)
            tik3 = ws.Cells(j, 9)
        End If
        ws.Cells(4, 17) = Greatest_TV
        ws.Cells(4, 16) = tik3
        
    Next j
        
    
    For k = 2 To LastRow
        If ws.Cells(k, 1) = ws.Cells(k + 1, 1) Then
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(k + 1, 7)
        Else
            ws.Cells(TSV_Counter, 12) = Total_Stock_Volume
            TSV_Counter = TSV_Counter + 1
            Total_Stock_Volume = 0
        End If
    Next k
            
    
    
    Next ws
    
    
    
End Sub

