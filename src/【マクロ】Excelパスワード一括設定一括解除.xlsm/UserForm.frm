VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3396
   ClientLeft      =   -300
   ClientTop       =   -996
   ClientWidth     =   4980
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'一括設定処理
Private Sub LockButton_Click()
  '
End Sub

'一括解除処理
Private Sub unlockButton_Click()
  'アラートの非表示
  Application.DisplayAlerts = False
  
  Set objFileSys = CreateObject("Scripting.FileSystemObject")
  
  'カレントディレクトリのファイル名を取得
  Dim FileName As String
  FileName = Dir(ThisWorkbook.Path & "\*")

  'ログ出力領域の初期化
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ws.Range("B:B").Clear
  
  'ログ出力開始位置の初期化
  Dim logRowIndex As Integer
  logRowIndex = 2
  
  'パスワード
  Dim password As String
  password = TextBox1.Value
  
  '解除ファイル出力先の設定
  Dim savePath As String
  savePath = ThisWorkbook.Path & "\解除後"
  If Not objFileSys.FolderExists(savePath) Then objFileSys.CreateFolder savePath

  '処理開始
  Do While FileName <> ""
    
    '処理対象外ファイルをスキップする。
    If FileName = ThisWorkbook.Name Or _
    (Right(FileName, 5) <> ".xlsx" And Right(FileName, 5) <> ".xlsm" And Right(FileName, 4) <> ".xls") Then
      Call OutputLog(ws, FileName & " skipped.", logRowIndex)
      GoTo Continue
    End If
    
    'ファイルを開く
    On Error Resume Next
    Dim wb As Workbook    
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & FileName, Password:=password, WriteResPassword:=password)
    
    'ファイルオープンエラーの場合、対象ファイルをスキップする。
    If Err.Number <> 0 Then
      Call OutputLog(ws, FileName & " skipped. " & Err.Description, logRowIndex)
      Call CloseFile(wb)
      GoTo Continue
    End If
    
    '対象ファイルが保護されている場合解除する。
    If wb.ProtectWindows = True Then
      wb.Unprotect Password:=password
    End If
    
    'パスワードを解除して上書きして閉じる。
    wb.SaveAs savePath & "\確認用_" & FileName, Password:=""
    Call OutputLog(ws, FileName & " unlocked.", logRowIndex)
    Call CloseFile(wb)
  
Continue:
    
    '次の対象ファイルを処理する。
    FileName = Dir()
  Loop

  MsgBox "Finish."
 
  Set objFileSys = Nothing
  Application.DisplayAlerts = True
End Sub

'Excelファイルを閉じる
' wb:指定のExcel
Function CloseFile(ByRef wb As Workbook)
  wb.Close
  wb = Nothing
End Function

'ログを出力する。
' ws:出力対象のWorksheet
' message:ログ内容
' rowNumber:ログの出力位置の行番号
Function OutputLog(ByRef ws As Worksheet, ByVal message As String, ByRef rowNumber As Integer)
  ws.Cells(rowNumber, 2).Value = message  
  rowNumber = rowNumber + 1
End Function