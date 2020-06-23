Attribute VB_Name = "Module1"
Sub The_VBA_Of_Wall_Street()

Dim Activews As Worksheet

For Each Activews In Worksheets
 
 Dim Ticker_Symbol As String
 Dim Stock_Volume_Total As Double
 Stock_Volume_Total = 0
 Dim Open_Price As Double
 Open_Price = 0
 Dim Close_Price As Double
 Close_Price = 0
 Dim Yearly_Change As Double
 Yearly_Change = 0
 Dim Percentage_Change As Double
 Percentage_Change = 0

 Dim Max_Ticker_Symbol As String
 Max_Ticker_Symbol = " "
 Dim Min_Ticker_Symbol As String
 Min_Ticker_Symbol = " "
 Dim Max_Increase_Percent As Double
 Max_Increase_Percent = 0
 Dim Min_Increase_Percent As Double
 Min_Increase_Percent = 0
 Dim Max_Volume_Ticker As String
 Max_Volume_Ticker = " "
 Dim Max_Stock_Volume As Double
 Max_Stock_Volume = 0
 
 
 Dim Summary_Table_Row As Double
 Summary_Table_Row = 2
 
 Dim lastRow As Long
 Dim i As Long
 
 lastRow = Activews.Cells(Rows.Count, 1).End(xlUp).Row
 
 Activews.Range("I1").Value = "Ticket Symbol"
 Activews.Range("J1").Value = "Yearly Change"
 Activews.Range("K1").Value = "Percent Change"
 Activews.Range("L1").Value = "Total Stock Volume"
  
 Activews.Range("O2").Value = "Greatest % Increase"
 Activews.Range("O3").Value = "Greatest % Decrease"
 Activews.Range("O4").Value = "Greatest Total Volume"
 Activews.Range("P1").Value = "Ticker"
 Activews.Range("Q1").Value = "Value"
 
 Open_Price = Activews.Cells(2, 3).Value
 
 For i = 2 To lastRow

  If Activews.Cells(i + 1, 1).Value <> Activews.Cells(i, 1).Value Then
    
    Ticker_Symbol = Activews.Cells(i, 1).Value
    
    Close_Price = Activews.Cells(i, 6).Value
    
    Yearly_Change = Close_Price - Open_Price
    If Open_Price <> 0 Then
     Percentage_Change = (Yearly_Change / Open_Price) * 100
    
    End If
    
    Stock_Volume_Total = Stock_Volume_Total + Activews.Cells(i, 7).Value
    
    Activews.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
    Activews.Range("J" & Summary_Table_Row).Value = Yearly_Change
    If (Yearly_Change > 0) Then
     Activews.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    ElseIf (Yearly_Change <= 0) Then
     Activews.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
    
    Activews.Range("K" & Summary_Table_Row).Value = (CStr(Percentage_Change) & "%")
    Activews.Range("L" & Summary_Table_Row).Value = Stock_Volume_Total
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    Yearly_Change = 0
    Close_Price = 0
    
    Open_Price = Activews.Cells(i + 1, 3).Value
    
    If (Percentage_Change > Max_Increase_Percent) Then
     Max_Increase_Percent = Percentage_Change
     Max_Ticker_Symbol = Ticker_Symbol
    
    ElseIf (Percentage_Change < Min_Increase_Percent) Then
     Min_Increase_Percent = Percentage_Change
     Min_Ticker_Symbol = Ticker_Symbol
    End If
    
    If (Stock_Volume_Total > Max_Stock_Volume) Then
     Max_Stock_Volume = Stock_Volume_Total
     Max_Volume_Ticker = Ticker_Symbol
    End If
    
    Percentage_Change = 0
    Stock_Volume_Total = 0
    
  Else
       
    Stock_Volume_Total = Stock_Volume_Total + Activews.Cells(i + 1, 7).Value
        
  End If

 Next i
    
    Activews.Range("Q2").Value = (CStr(Max_Increase_Percent) & "%")
    Activews.Range("Q3").Value = (CStr(Min_Increase_Percent) & "%")
    Activews.Range("Q4").Value = Max_Stock_Volume
    Activews.Range("P2").Value = Max_Ticker_Symbol
    Activews.Range("P3").Value = Min_Ticker_Symbol
    Activews.Range("P4").Value = Max_Volume_Ticker
    
    Activews.Columns("I:Q").AutoFit
    
Next Activews
End Sub
