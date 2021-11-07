Sub Ticker():

Dim Ws As Worksheet
Dim Title_Summary_Table As Boolean
Dim Summary_Data As Boolean

Title_Summary_Table = False
Summary_Data = True

'Loop for all worksheets

For Each Ws In Worksheets
    
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim open_price As Double
    Dim close_price As Double
    
    
    Ticker = " "
    Total_Stock_Volume = 0
    Yearly_Change = 0
    Percent_Change = 0
    open_price = 0
    close_price = 0
    
    
'Summary Table

    Dim Min_Ticker As String
    Dim Max_Ticker As String
    Dim Max_Vol_Ticker As String
    Dim Min_Percent As Double
    Dim Max_Percent As Double
    Dim Max_Volume As Double

    Min_Ticker = " "
    Max_Ticker = " "
    Max_Vol_Ticker = " "
    Min_Percent = 0
    Max_Percent = 0
    Max_Volume = 0

    Dim lastrow As Long
    Dim Ticker_Summary As Long
    Dim i As Long
    
    lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
    Ticker_Summary = 2
    
    If Title_Summary_Table Then
        Ws.Range("I1").Value = "Ticker"
        Ws.Range("J1").Value = "Yearly Change"
        Ws.Range("K1").Value = "Percent Change"
        Ws.Range("L1").Value = "Total Stock Volume"
        Ws.Range("P1").Value = "Ticker"
        Ws.Range("Q1").Value = "Value"
        Ws.Range("O2").Value = "Greatest % Increase"
        Ws.Range("O3").Value = "Greatest % Decrease"
        Ws.Range("O4").Value = "Greatest Total Volume"
        
        Else
        Title_Summary_Table = True
    End If
    
'Transferring onto the summary table

    open_price = Ws.Cells(2, 3).Value

    For i = 2 To lastrow
    

        If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            Ticker = Ws.Cells(i, 1).Value
            
            close_price = Ws.Cells(i, 6).Value
            Yearly_Change = close_price - open_price
            
            If open_price <> 0 Then
                Percent_Change = (Yearly_Change / open_price) * 100
            
        End If
            
        Total_Stock_Volume = Total_Stock_Volume + Ws.Cells(i, 7).Value
        
        
        Ws.Range("I" & Ticker_Summary).Value = Ticker
        Ws.Range("J" & Ticker_Summary).Value = Yearly_Change
        Ws.Range("K" & Ticker_Summary).Value = (CStr(Percent_Change) & "%")
        Ws.Range("L" & Ticker_Summary).Value = Total_Stock_Volume

'Conditional Formatting
            
        If (Percent_Change > 0) Then
            Ws.Range("J" & Ticker_Summary).Interior.ColorIndex = 4
            
            ElseIf (Percent_Change <= 0) Then
            Ws.Range("J" & Ticker_Summary).Interior.ColorIndex = 3
            
            End If
        
        Ticker_Summary = Ticker_Summary + 1
        Yearly_Change = 0
        Percent_Change = 0
        close_price = 0
        open_price = Ws.Cells(i + 1, 3).Value
        
        If (Percent_Change > Max_Percent) Then
            Max_Percent = Percent_Change
            Max_Ticker = Ticker
            
        ElseIf (Percent_Change < Min_Percent) Then
            Min_Percent = Percent_Change
            Min_Ticker = Ticker
        End If
        
        If (Total_Stock_Volume > Max_Volume) Then
            Max_Volume = Total_Stock_Volume
            Max_Vol_Ticker = Ticker
        End If
        
        Percent_Change = 0
        Total_Stock_Volume = 0
                
        Else
        
        Total_Stock_Volume = Total_Stock_Volume + Ws.Cells(i, 7).Value
        
        End If


    Next i
    
    If Not Summary_Data Then
    
        Ws.Range("P2").Value = Max_Ticker
        Ws.Range("P3").Value = Min_Ticker
        Ws.Range("P4").Value = Max_Vol_Ticker
        Ws.Range("Q2").Value = (CStr(Max_Percent) & "%")
        Ws.Range("Q3").Value = (CStr(Min_Percent) & "%")
        Ws.Range("Q4").Value = Max_Volume
    
    Else
        Summary_Data = False
    End If
    
  Next Ws

End Sub