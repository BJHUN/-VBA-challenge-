Attribute VB_Name = "Module1"
Sub Stock_Analysis()

Dim Ticker As String
Dim Total_Ticker As Double
Total_Ticker = 0
Dim Open_Price As Double
Open_Price = 0
Dim Close_Price As Double
Close_Price = 0
Dim Delta_Price As Double
Delta_Price = 0
Dim Delta_Percent As Double
Delta_Percent = 0
Dim Summary_Table As Long
Summary_Table = 2
Dim Lastrow As Long
Dim L As Long
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Open_Price = Cells(2, 3).Value
For L = 2 To Lastrow

If Cells(L + 1, 1).Value <> Cells(L, 1).Value Then

Ticker = Cells(L, 1).Value
Close_Price = Cells(L, 6).Value
Delta_Price = Close_Price - Open_Price
If Open_Price <> 0 Then
Delta_Percent = (Delta_Price / Open_Price) * 100
End If
Total_Ticker = Total_Ticker + Cells(L, 7).Value
Range("I" & Summary_Table).Value = Ticker
Range("J" & Summary_Table).Value = Delta_Price
Range("K" & Summary_Table).Value = (CStr(Delta_Percent) & "%")
Range("L" & Summary_Table).Value = Total_Ticker

If (Delta_Price > 0) Then
Range("J" & Summary_Table).Interior.ColorIndex = 4
ElseIf (Delta_Price <= 0) Then
Range("J" & Summary_Table).Interior.ColorIndex = 3
End If

Summary_Table = Summary_Table + 1
Delta_Price = 0
Close_Price = 0
Open_Price = Cells(L + 1, 3).Value

Else
Total_Ticker = Total_Ticker + Cells(L, 7).Value

End If
Next L

End Sub

