Option Explicit
' Run only in the owned qualification controller beside embedded-source.xlsx.
' No automatic entry points, external links, or application-setting changes.
Public Sub QualifyNativeEmbedding()
    Dim wb As Object
    Dim ws As Object
    Dim embedded As Object
    Dim root As String
    Dim optionError As Long
    Dim optionMessage As String
    Dim failureNumber As Long
    Dim failureMessage As String
    On Error GoTo Failed
#If Mac Then
    root = ThisWorkbook.Path & "/"
#Else
    root = ThisWorkbook.Path & "\"
#End If
    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Native"
    ws.Cells(1, 1).Value2 = "Native encrypted embedding control"
    ws.Cells(1, 2).Value2 = 42
    ws.Cells(2, 2).Value2 = 8
    ws.Cells(3, 2).Formula = "=SUM(B1:B2)"
    wb.BuiltinDocumentProperties("Title").Value = "Native encrypted auxiliary control"
    wb.BuiltinDocumentProperties("Subject").Value = "Native encrypted nested storage"
    Set embedded = ws.OLEObjects.Add(Filename:=root & "embedded-source.xlsx", Link:=False, DisplayAsIcon:=True)
    ws.Cells(4, 2).Value2 = ws.OLEObjects.Count
    wb.SaveAs Filename:=root & "excel-native-embedding.plain.xls", FileFormat:=56
    On Error Resume Next
    wb.SetPasswordEncryptionOptions PasswordEncryptionFileProperties:=True
    optionError = Err.Number
    optionMessage = Err.Description
    Err.Clear
    On Error GoTo Failed
    ws.Cells(1, 4).Value2 = optionError
    ws.Cells(2, 4).Value2 = optionMessage
    wb.SaveAs Filename:=root & "excel-native-embedding.encrypted.xls", FileFormat:=56, Password:="test-password"
    wb.Close SaveChanges:=False
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureMessage = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Worksheets(1).Cells(1, 26).Value2 = "Failure " & failureNumber & ": " & failureMessage
    On Error GoTo 0
End Sub
