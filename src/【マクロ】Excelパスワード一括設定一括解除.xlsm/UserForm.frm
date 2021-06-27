VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2736
   ClientLeft      =   -468
   ClientTop       =   -1632
   ClientWidth     =   5232
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'�ꊇ�ݒ菈��
Private Sub LockButton_Click()
  '�A���[�g�̔�\��
  Application.DisplayAlerts = False
  
  Set objFileSys = CreateObject("Scripting.FileSystemObject")
  
  '�J�����g�f�B���N�g���̃t�@�C�������擾
  Dim FileName As String
  FileName = Dir(ThisWorkbook.Path & "\*")

  '���O�o�͗̈�̏�����
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ws.Range("B:B").Clear
  
  '���O�o�͊J�n�ʒu�̏�����
  Dim logRowIndex As Integer
  logRowIndex = 2
  
  '�p�X���[�h
  Dim password As String
  password = TextBox1.Value
  
  '�p�X�t�^�t�@�C���o�͐�̐ݒ�
  Dim savePath As String
  savePath = ThisWorkbook.Path & "\lock"
  If Not objFileSys.FolderExists(savePath) Then objFileSys.CreateFolder savePath

  '�����J�n
  Do While FileName <> ""
    
    '�����ΏۊO�t�@�C�����X�L�b�v����B
    If FileName = ThisWorkbook.Name Or _
    (Right(FileName, 5) <> ".xlsx" And Right(FileName, 4) <> ".xls") Then
      Call OutputLog(ws, FileName & " skipped.", logRowIndex)
      GoTo Continue
    End If
    
    '�t�@�C�����J��
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & FileName, WriteResPassword:=password)
    
    '�t�@�C���I�[�v���G���[�̏ꍇ�A�Ώۃt�@�C�����X�L�b�v����B
    If Err.Number <> 0 Then
      Call OutputLog(ws, FileName & " skipped. " & Err.Description, logRowIndex)
      Call CloseFile(wb)
      GoTo Continue
    End If
    
    '�Ώۃt�@�C�����ی삳��Ă���ꍇ��������B
    If wb.ProtectWindows = True Then
      wb.Unprotect password:=password
    End If
    
    '�p�X���[�h���������ď㏑�����ĕ���B
    wb.SaveAs savePath & "\" & FileName, password:=password
    Call OutputLog(ws, FileName & " locked.", logRowIndex)
    Call CloseFile(wb)
  
Continue:
    
    '���̑Ώۃt�@�C������������B
    FileName = Dir()
  Loop

  MsgBox "Finish."
 
  Set objFileSys = Nothing
  Application.DisplayAlerts = True
End Sub

'�ꊇ��������
Private Sub unlockButton_Click()
  '�A���[�g�̔�\��
  Application.DisplayAlerts = False
  
  Set objFileSys = CreateObject("Scripting.FileSystemObject")
  
  '�J�����g�f�B���N�g���̃t�@�C�������擾
  Dim FileName As String
  FileName = Dir(ThisWorkbook.Path & "\*")

  '���O�o�͗̈�̏�����
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ws.Range("B:B").Clear
  
  '���O�o�͊J�n�ʒu�̏�����
  Dim logRowIndex As Integer
  logRowIndex = 2
  
  '�p�X���[�h
  Dim password As String
  password = TextBox1.Value
  
  '�����t�@�C���o�͐�̐ݒ�
  Dim savePath As String
  savePath = ThisWorkbook.Path & "\unlock"
  If Not objFileSys.FolderExists(savePath) Then objFileSys.CreateFolder savePath

  '�����J�n
  Do While FileName <> ""
    
    '�����ΏۊO�t�@�C�����X�L�b�v����B
    If FileName = ThisWorkbook.Name Or _
    (Right(FileName, 5) <> ".xlsx" And Right(FileName, 5) <> ".xlsm" And Right(FileName, 4) <> ".xls") Then
      Call OutputLog(ws, FileName & " skipped.", logRowIndex)
      GoTo Continue
    End If
    
    '�t�@�C�����J��
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & FileName, password:=password, WriteResPassword:=password)
    
    '�t�@�C���I�[�v���G���[�̏ꍇ�A�Ώۃt�@�C�����X�L�b�v����B
    If Err.Number <> 0 Then
      Call OutputLog(ws, FileName & " skipped. " & Err.Description, logRowIndex)
      Call CloseFile(wb)
      GoTo Continue
    End If
    
    '�Ώۃt�@�C�����ی삳��Ă���ꍇ��������B
    If wb.ProtectWindows = True Then
      wb.Unprotect password:=password
    End If
    
    '�p�X���[�h���������ď㏑�����ĕ���B
    wb.SaveAs savePath & "\sumiyoshi_" & FileName, password:=""
    Call OutputLog(ws, FileName & " unlocked.", logRowIndex)
    Call CloseFile(wb)
  
Continue:
    
    '���̑Ώۃt�@�C������������B
    FileName = Dir()
  Loop

  MsgBox "Finish."
 
  Set objFileSys = Nothing
  Application.DisplayAlerts = True
End Sub

'Excel�t�@�C�������
' wb:�w���Excel
Function CloseFile(ByRef wb As Workbook)
  wb.Close
  wb = Nothing
End Function

'���O���o�͂���B
' ws:�o�͑Ώۂ�Worksheet
' message:���O���e
' rowNumber:���O�̏o�͈ʒu�̍s�ԍ�
Function OutputLog(ByRef ws As Worksheet, ByVal message As String, ByRef rowNumber As Integer)
  ws.Cells(rowNumber, 2).Value = message
  rowNumber = rowNumber + 1
End Function
