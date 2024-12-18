Attribute VB_Name = "Main"

Sub ExportPane()
  Dim EXPORTS() As config.Export
  EXPORTS = config.config()
  Util.ExportToWord EXPORTS
  MsgBox "Successuflly exported all data", vbInformation, "Export"
End Sub
