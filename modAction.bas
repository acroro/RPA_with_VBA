Attribute VB_Name = "modAction"
'--- modActions ---
Option Explicit

Private gTargetBook As Workbook
Private gCopyRange As Range

Public Sub Action_OpenBook(ByVal filePath As String)
  If filePath = "" Then Err.Raise vbObjectError + 1, , "OpenBook: filePath is empty"
  Set gTargetBook = Workbooks.Open(filePath)
End Sub

Public Sub Action_CopyRange(ByVal target As String)
  Dim wsName As String, addr As String
  SplitSheetAddress target, wsName, addr

  If gTargetBook Is Nothing Then Err.Raise vbObjectError + 2, , "CopyRange: target book not opened"
  Set gCopyRange = gTargetBook.Worksheets(wsName).Range(addr)
End Sub

Public Sub Action_PasteRange(ByVal target As String)
  Dim wsName As String, addr As String
  SplitSheetAddress target, wsName, addr

  If gTargetBook Is Nothing Then Err.Raise vbObjectError + 3, , "PasteRange: target book not opened"
  If gCopyRange Is Nothing Then Err.Raise vbObjectError + 4, , "PasteRange: nothing copied"

  With gTargetBook.Worksheets(wsName).Range(addr)
    .Resize(gCopyRange.Rows.count, gCopyRange.Columns.count).Value = gCopyRange.Value
  End With
End Sub

Public Sub Action_SaveBook()
  If Not gTargetBook Is Nothing Then gTargetBook.Save
End Sub

Private Sub SplitSheetAddress(ByVal s As String, ByRef wsName As String, ByRef addr As String)
  Dim p() As String
  p = Split(s, "!")
  If UBound(p) <> 1 Then Err.Raise vbObjectError + 10, , "Invalid target format (use Sheet!A1): " & s
  wsName = p(0)
  addr = p(1)
End Sub
