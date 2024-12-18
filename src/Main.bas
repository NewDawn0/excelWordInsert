Attribute VB_Name = "Main"

Sub ExportPane()
  Dim EXPORTS() As config.Export
  EXPORTS = config.config()
  MsgBox "Successuflly exported to <test>", vbInformation, "Export"
End Sub
