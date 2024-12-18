Attribute VB_NAME = "Util"

Sub ExportToWord(ByRef Exports() As config.Export)
  CheckData Exports
End Sub

Sub CheckData(ByRef Exports() As config.Export)
  Dim i As Long
  Dim ws As Worksheet
  Dim paneExists As Boolean
  
  For i = LBound(Exports) To UBound(Exports)
    paneExists = False
    For Each ws In ThisWorkbook.Worksheets
      If ws.Name = Exports(i).pane Then
        paneExists = True
        Exit For
      End If
    Next ws
    
    If Not paneExists Then
      MsgBox "The pane '" & Exports(i).pane & "' does not exist in this Workbook", vbCritical, "Pane not found"
      End ' Abort Since the data does not exist '
    End If
  Next i
End Sub
