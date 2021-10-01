Attribute VB_Name = "Module1"
Sub HW()
Attribute HW.VB_ProcData.VB_Invoke_Func = " \n14"
For j = 1 To Worksheets.Count
Worksheets(j).Activate
Dim totalvol As Double
Dim counter As Long
Dim Openvl As Double
Dim closevl As Double


LastRow = Cells(Rows.Count, 1).End(xlUp).Row
counter = 2
Openvl = Cells(2, 3).Value
Cells(1, 9).Value = "Ticker"
Cells(1, 11).Value = "Precent change"
Cells(1, 10).Value = "Yearly change"
Cells(1, 12).Value = "Totalvol "
Cells(2, 9).Value = Cells(2, 1).Value

For i = 2 To LastRow

If Cells(i, 1).Value = Cells(i + 1, 1).Value Then

totalvol = Cells(i, 7).Value + totalvol

Else
 
 totalvol = Cells(i, 7).Value + totalvol
 Cells(counter, 12).Value = totalvol
 Cells(counter + 1, 9).Value = Cells(i + 1, 1).Value
 
 closevl = Cells(i, 6).Value
 
 Cells(counter, 10).Value = closevl - Openvl

 
 If Openvl = 0 Then
 Cells(counter, 11).Value = "0%"
Else
Cells(counter, 11).Value = ((closevl - Openvl) / Openvl) * 100 & "%"
 End If
If Cells(counter, 10).Value >= 0 Then

Cells(counter, 10).Interior.Color = RGB(0, 255, 0)
Else
Cells(counter, 10).Interior.Color = RGB(255, 0, 0)

End If
totalvol = 0
counter = counter + 1
Openvl = Cells(i + 1, 3).Value
 End If
 Next i
 Next j
  
 
 
End Sub
