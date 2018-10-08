Attribute VB_Name = "Module1"
Option Explicit



Function Find1(WS As Worksheet, column As Long, data As Variant) As Range
  Dim ret As Range
  
  Dim row As Long
  For row = 1 To WS.UsedRange.Rows.Count
    If CDate(data) = WS.Cells(row, 1).Value Then
      Set ret = Cells(row, 1)
      Exit For
    End If
  Next
  
  Set Find1 = ret
End Function



Function test()
  Dim R As Range
  
  Dim column As Long
  
  For column = 1 To 4
    Set R = Find1(ActiveSheet, column, "2010/1/20")
    Debug.Print "Column" & column & ":" & R.Address
  Next
  
End Function
