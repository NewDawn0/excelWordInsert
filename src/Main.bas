Attribute VB_Name = "Main"

Sub ExportData()
  ' Save documents reminder '
  Dim resp As VbMsgBoxResult
  resp = MsgBox("This macro will close all your word documents. " & vbCrLf & "Make sure all your word documents saved.", vbOKCancel + vbInformation, "Do you want to proceed?")
  If resp = vbCancel Then
    End ' User aborted '
  End If
  Dim EXPORTS() As config.Export
  EXPORTS = config.config()
  Util.ExportToWord EXPORTS
  MsgBox "Successuflly exported all data", vbInformation, "Export"
End Sub