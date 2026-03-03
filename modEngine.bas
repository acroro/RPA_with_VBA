Attribute VB_Name = "modEngine"
'--- modEngine ---
Option Explicit

Public Type Command
  StepNo As Long
  Action As String
  target As String
  Value As String
  Condition As String
  Retry As Long
End Type

'========================
' エントリーポイント
'========================
Public Sub RunCommands()
  Dim wsCmd As Worksheet
  Set wsCmd = GetCommandsSheet()

  Dim cmds() As Command
  cmds = LoadCommandsFromSheet(wsCmd)

  Dim i As Long
  For i = LBound(cmds) To UBound(cmds)

    If EvaluateCondition(cmds(i).Condition) Then
      If Not ExecuteWithRetry(cmds(i)) Then
        LogError "Step " & cmds(i).StepNo & " failed. Action=" & cmds(i).Action
        Exit For
      End If
    Else
      LogInfo "Step " & cmds(i).StepNo & " skipped by condition."
    End If

  Next i

  LogInfo "RunCommands finished."
End Sub

'========================
' Commandsシート取得（名前揺れ吸収）
'========================
Private Function GetCommandsSheet() As Worksheet
  On Error Resume Next
  Set GetCommandsSheet = ThisWorkbook.Worksheets("Commands")
  On Error GoTo 0
  If Not GetCommandsSheet Is Nothing Then Exit Function

  On Error Resume Next
  Set GetCommandsSheet = ThisWorkbook.Worksheets("指示シート")
  On Error GoTo 0
  If Not GetCommandsSheet Is Nothing Then Exit Function

  '最後の保険：アクティブシート
  Set GetCommandsSheet = ActiveSheet
End Function

'========================
' 指示読み込み
'========================
Private Function LoadCommandsFromSheet(ByVal ws As Worksheet) As Command()
  Dim lastRow As Long
  lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

  Dim tmp() As Command
  Dim count As Long: count = 0

  Dim r As Long
  For r = 2 To lastRow '1行目ヘッダ想定
    If Trim$(ws.Cells(r, "B").Value) = "" Then Exit For

    count = count + 1
    ReDim Preserve tmp(1 To count)

    tmp(count).StepNo = val(ws.Cells(r, "A").Value)
    If tmp(count).StepNo = 0 Then tmp(count).StepNo = r - 1 'Aが空なら行番号基準

    tmp(count).Action = Trim$(ws.Cells(r, "B").Value)
    tmp(count).target = Trim$(ws.Cells(r, "C").Value)
    tmp(count).Value = Trim$(ws.Cells(r, "D").Value)
    tmp(count).Condition = Trim$(ws.Cells(r, "E").Value)
    tmp(count).Retry = val(ws.Cells(r, "F").Value)
  Next r

  If count = 0 Then
    ReDim tmp(1 To 1)
  End If

  LoadCommandsFromSheet = tmp
End Function

'========================
' 条件評価（空ならTrue）
' 例：Condition に 1=1 や A1>0 のような式
'========================
Private Function EvaluateCondition(ByVal expr As String) As Boolean
  expr = Trim$(expr)
  If expr = "" Then
    EvaluateCondition = True
    Exit Function
  End If

  On Error GoTo EH
  EvaluateCondition = CBool(Application.Evaluate(expr))
  Exit Function

EH:
  '式が評価できない場合は安全側：False（必要ならTrueに変更）
  EvaluateCondition = False
  LogError "Condition evaluate error: " & expr
End Function

'========================
' リトライ実行
'========================
Private Function ExecuteWithRetry(ByRef cmd As Command) As Boolean
  Dim n As Long, maxTry As Long
  maxTry = cmd.Retry + 1
  If maxTry < 1 Then maxTry = 1

  For n = 1 To maxTry
    If ExecuteCommand(cmd.Action, cmd.target, cmd.Value) Then
      ExecuteWithRetry = True
      Exit Function
    End If
    LogInfo "Retry " & n & "/" & maxTry & " : Step " & cmd.StepNo
    DoEvents
  Next n

  ExecuteWithRetry = False
End Function

'========================
' Actionディスパッチ（ここを増やすとRPAが育つ）
'========================
Private Function ExecuteCommand(ByVal act As String, ByVal tgt As String, ByVal val As String) As Boolean
  On Error GoTo EH

  Select Case UCase$(Trim$(act))
    Case "OPENBOOK"
      Action_OpenBook val
    Case "COPYRANGE"
      Action_CopyRange tgt
    Case "PASTERANGE"
      Action_PasteRange tgt
    Case "SAVEBOOK"
      Action_SaveBook
    Case "END"
      ExecuteCommand = True
      Exit Function
    Case Else
      LogError "Unknown action: " & act
      ExecuteCommand = False
      Exit Function
  End Select

  ExecuteCommand = True
  Exit Function

EH:
  LogError "ExecuteCommand error: act=" & act & " tgt=" & tgt & " val=" & val & " / " & Err.Description
  ExecuteCommand = False
End Function

'========================
' ログ（シート"Log"に追記）
'========================
Private Sub LogInfo(ByVal msg As String)
  AppendLog "INFO", msg
End Sub

Private Sub LogError(ByVal msg As String)
  AppendLog "ERROR", msg
End Sub

Private Sub AppendLog(ByVal level As String, ByVal msg As String)
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("Log")
  On Error GoTo 0

  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = "Log"
    ws.Range("A1:C1").Value = Array("Time", "Level", "Message")
  End If

  Dim r As Long
  r = ws.Cells(ws.Rows.count, "A").End(xlUp).Row + 1
  ws.Cells(r, "A").Value = Now
  ws.Cells(r, "B").Value = level
  ws.Cells(r, "C").Value = msg
End Sub



