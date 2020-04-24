VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub vbaA():

Dim total As Double

firstRow = Cells(Rows.Count, "A").End(xlUp).Row

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"

For i = 2 To firstRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Range("I" & 2 + j).Value = Cells(i, 1).Value
    Range("J" & 2 + j).Value = total
    
    j = j + 1
    
Else
    total = total + Cells(i, 7).Value
End If

Next i

End Sub



