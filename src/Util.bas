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

      ' Check if the resolved file exists
    If Dir(resolvedPath) = "" Then
      MsgBox "The file '" & resolvedPath & "' does not exist.", vbCritical, "File Not Found"
      End ' Abort since the file is missing '
    End If

    
    If Not paneExists Then
      MsgBox "The pane '" & Exports(i).pane & "' does not exist in this Workbook", vbCritical, "Pane not found"
      End ' Abort Since the data does not exist '
    End If
  Next i
End Sub
