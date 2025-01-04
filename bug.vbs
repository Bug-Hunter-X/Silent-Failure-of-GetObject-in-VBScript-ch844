Function GetObject(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Set obj = CreateObject(progID)
  End If
  Set GetObject = obj
End Function

Sub Main()
  Dim objExcel
  Set objExcel = GetObject("Excel.Application")
  If objExcel Is Nothing Then
    MsgBox "Excel is not running."
    WScript.Quit
  End If

  ' ... rest of your code to work with Excel ...
End Sub