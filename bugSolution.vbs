Function GetObjectSafe(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear ' Clear the error object
    Set obj = Nothing
  End If
  Set GetObjectSafe = obj
End Function

Sub Main()
  Dim objExcel
  Set objExcel = GetObjectSafe("Excel.Application")
  If objExcel Is Nothing Then
    MsgBox "Excel is not running or there was an error accessing it.", vbExclamation
    WScript.Quit
  End If

  ' ... rest of your code to work with Excel ...
  objExcel.Quit
  Set objExcel = Nothing
End Sub