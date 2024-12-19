Attribute VB_NAME = "Util"

Sub ExportToWord(ByRef Exports() As config.Export)
  CheckData Exports
End Sub

Sub CheckData(ByRef Exports() As config.Export)
  Dim i As Long
  Dim ws As Worksheet
  Dim paneExists As Boolean
  Dim filePath As String
  
  For i = LBound(Exports) To UBound(Exports)
    ' Check for pane '
    paneExists = False
    For Each ws In ThisWorkbook.Worksheets
      If ws.Name = Exports(i).pane Then
        paneExists = True
        Exit For
      End If
    Next ws
    ' Check for file '
    filePath = Exports(i).file
    If Left(filePath, 2) = "./" Then
      resolvedPath = ThisWorkbook.Path & "/" & Mid(filePath, 3)
    ElseIf Left(filePath, 2) = ".\" Then
      resolvedPath = ThisWorkbook.Path & "\" & Mid(filePath, 3)
    Else
      resolvedPath = filePath
    End If
    If Dir(resolvedPath) = "" Then
      MsgBox "The file '" & resolvedPath & "' does not exist.", vbCritical, "File Not Found"
      End ' Abort since the file is missing '
    End If
    ' Check for marker in file '
    If Not CheckMarkerExists(resolvedPath, Exports(i).marker) Then
      MsgBox "The marker '" & Exports(i).marker & "' was not found in the file '" & resolvedPath & "'.", vbCritical, "Marker Not Found"
      End ' Abort since the marker does not exist '
    End If
    
    If Not paneExists Then
      MsgBox "The pane '" & Exports(i).pane & "' does not exist in this Workbook", vbCritical, "Pane not found"
      End ' Abort since the data does not exist '
    End If
  Next i
End Sub

Function CheckMarkerExists(ByVal filePath As String, ByVal marker As String) As Boolean
  Dim wordApp As Object
  Dim doc As Object
  Dim markerFound As Boolean
  markerFound = False
  
  Set wordApp = GetObject(, "Word.Application")
  ' if wordApp Is Nothing Then
  '  Set wordApp = CreateObject("Word.Application")
  ' End If
  
  Set doc = wordApp.Documents.Open(filePath)
  wordApp.Visible = True
  
  With doc.Content.Find
    .Text = marker
    .MatchCase = True
    .MatchWholeWord = True
    .Execute
    markerFound = .Found
  End With
  doc.Close
  wordApp.Quit
  Set doc = Nothing
  Set wordApp = Nothing
  CheckMarkerExists = markerFound
End Function