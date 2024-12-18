Attribute VB_NAME = "Config"

Type Export
  startCell As String ' eg. "A1" '
  endCell As String ' eg. "B4" '
  marker As String ' Your marker in the word where the cells are copied to'
  pane As String ' Your excel pane name '
  ' Your .dotm file containing the marker to which the text is copied to'
  file As String  'eg. "C:\Path\To\TemplateFile.dotm" '
End Type

Function config() As Export()
  Dim EXPORTS(1 To 1) As Export 'In (1 To <number of exports>) for 2 exports it should be (1 To 2) '
  ' Exports defined here'
  EXPORTS(1).startCell = 1
  EXPORTS(1).endCell = 1
  EXPORTS(1).pane = "HolySheetItWorks"
  EXPORTS(1).marker = "Marker!"
  EXPORTS(1).file = "./test.docx"
  config = EXPORTS
End Function