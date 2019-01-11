Attribute VB_Name = "moduleIO"
' Text Scripting on VBA v1.0.0
' last update: 2019-01-09
' Modified by Rintaro Nagashima
' Original coded by HATANO Hirokazu
'
' Detail: http://rsh.csh.sh/text-scripting-vba/
'  See Also: http://d.hatena.ne.jp/language_and_engineering/20090731/p1

Option Explicit


'----------------------------- Consts ---------------

'���C�u�������X�g�̐ݒ� (�ݒu�t�H���_�̓��[�N�u�b�N�Ɠ����f�B���N�g��)
Const FILENAME_LIBLIST As String = "libdef.txt" '���C�u�������X�g�̃t�@�C����
Const FILEPATH_LIBLIST As String = "" '�G�N�Z���t�@�C�����猩�����C�u�������X�g�̑��΃p�X(���Ƀp�X��؂蕶���͓���Ȃ�����)

'���[�N�u�b�N �I�[�v�����Ɏ��s����(True) / ���Ȃ�(False)
Const ENABLE_WORKBOOK_OPEN As Boolean = False

'�V���[�g�J�b�g�L�[
Const SHORTKEY_RELOAD As String = "r" 'ctrl + r
Const SHORTKEY_EXPORT As String = "e" 'ctrl + e

'----------------------------- Workbook_open() ---------------

'���[�N�u�b�N �I�[�v�����Ɏ��s
Private Sub Workbook_Open()
  Call setShortKey
  If ENABLE_WORKBOOK_OPEN = True Then
    Call reloadModule
  End If
 End Sub

'���[�N�u�b�N �N���[�Y���Ɏ��s
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Call clearShortKey
 End Sub



'----------------------------- public Subs/Functions ---------------

Public Sub reloadModule()
Attribute reloadModule.VB_ProcData.VB_Invoke_Func = "r\n14"
  '�蓮�����[�h�p Public�֐�
  
  Dim msgError As String
  msgError = loadModule("." & application.PathSeparator & FILENAME_LIBLIST)
  
  If Len(msgError) > 0 Then
    MsgBox msgError
  End If
End Sub


Public Sub exportModules()
  '�蓮export�p Public�֐�
  Dim arrayModules As Variant
  Dim i As Integer
  Dim message As String
  Dim msgError As String
  Dim curPath As String

  '���W���[�����X�g�t�@�C���̑��݊m�F�A�ǂݍ��݁A�z��
  msgError = getModuleList(FILENAME_LIBLIST, arrayModules, curPath)
  If Len(msgError) > 0 Then GoTo ErrHandler
  '���W���[�����X�g
  For i = 0 To UBound(arrayModules) - 1
    arrayModules(i) = absPath(arrayModules(i), curPath)
  Next
  
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents

    '�G�N�X�|�[�g����t�@�C���̃t���p�X
    Dim pathModule As String
    pathModule = "" '������

    ' �g���q���w��
    Dim moduleType As String
    If component.Type = 1 Then moduleType = ".bas"
    If component.Type = 2 Then moduleType = ".cls"

    ' ���C�u�����̃t���p�X�̒��ɁA���W���[���Ɠ������O�̃t�@�C�������邩�`�F�b�N
    For i = 0 To UBound(arrayModules) - 1
      If InStr(arrayModules(i), application.PathSeparator & component.Name & moduleType) > 0 Then
        pathModule = arrayModules(i)
        Exit For
      End If
    Next i

    ' TODO ���W���[����������΃G���[�B�V�[�g�I�u�W�F�N�g�ƃ��W���[����؂蕪������@��������Ύ������Ăق���
    ' ���W���[�����G�N�X�|�[�g
    If pathModule <> "" Then exportModule component, pathModule, message
      
    Next

  ' ���������ꍇ�A�G�N�X�|�[�g�t�@�C���ꗗ��\��
  MsgBox message
  Exit Sub

ErrHandler:
  MsgBox msgError
End Sub




'----------------------------- main Subs/Functions ---------------

Private Function loadModule(ByVal pathConf As String) As String
  'Main: ���W���[�����X�g�t�@�C���ɏ����Ă���O�����C�u������ǂݍ��ށB

  '1. �S���W���[�����폜
  Dim isClear As Boolean
  isClear = clearModules
  
  If isClear = False Then
    loadModule = "Error: �W�����W���[���̑S�폜�Ɏ��s���܂����B"
    Exit Function
  End If
  
  
  '2. ���W���[�����X�g�t�@�C���̑��݊m�F
  '3. ���W���[�����X�g�t�@�C���̓ǂݍ���&�z��
  Dim arrayModules As Variant
  Dim msgError As String
  Dim curPath As String
  msgError = getModuleList(pathConf, arrayModules, curPath)
  If msgError <> "" Then GoTo msgErr
  
  '4. �e���W���[���t�@�C���ǂݍ���
  Dim i As Integer
  
  ' �z���0�n�܂�B(�ő�l: �z���-1)
  For i = 0 To UBound(arrayModules) - 1
    Dim pathModule As String
    pathModule = arrayModules(i)
    
    '4.1. ���W���[�����X�g�t�@�C���̑��݊m�F
    ' 4.1.1. ���W���[�����X�g�t�@�C���̐�΃p�X���擾
    pathModule = absPath(pathModule, curPath)
  
    ' 4.1.2. ���݃`�F�b�N
    Dim isExistModule As Boolean
    isExistModule = checkExistFile(pathModule)
  
    '4.2. ���W���[���ǂݍ���
    If isExistModule = True Then
      ThisWorkbook.VBProject.VBComponents.Import pathModule
    Else
      msgError = msgError & pathModule & " �͑��݂��܂���B" & vbcrlf
    End If
  Next i
  
msgErr:
  loadModule = msgError

End Function



'----------------------------- Functions / Subs ---------------

Private Sub exportModule(ByVal component As Object, ByVal pathModule As String, ByRef message As String)
  
  component.Export pathModule
  message = message & component.Name & " �� " & pathModule & " �Ƃ��ĕۑ����܂����B" & vbcrlf
  
End Sub




'----------------------------- common Functions / Subs ---------------
Private Function clearModules() As Boolean
  '�W�����W���[��/�N���X���W���[��������(�S�폜)
  
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents

    ' ���̃��W���[�����g�͍폜���Ȃ�
    If component.Name = "moduleIO" Then GoTo Continue

    '�W�����W���[��(Type=1) / �N���X���W���[��(Type=2)��S�č폜
    If component.Type = 1 Or component.Type = 2 Then
      ThisWorkbook.VBProject.VBComponents.Remove component
    End If
    
Continue:
  Next component
  
  '�W�����W���[��/�N���X���W���[���̍��v����1(���̃��W���[�����g�̂�)�ł����OK
  Dim cntBAS As Long
  cntBAS = countBAS()
  
  Dim cntClass As Long
  cntClass = countClasses()
        
  If cntBAS = 1 And cntClass = 0 Then
    clearModules = True
  Else
    clearModules = False
  End If

End Function



Private Function countBAS() As Long
  Dim count As Long
  count = countComponents(1) 'Type 1: bas
  countBAS = count
End Function



Private Function countClasses() As Long
  Dim count As Long
  count = countComponents(2) 'Type 2: class
  countClasses = count
End Function



Private Function countComponents(ByVal numType As Integer) As Long
  '���݂���W�����W���[��/�N���X���W���[���̐��𐔂���
  
  Dim i As Long
  Dim count As Long
  count = 0
  
  With ThisWorkbook.VBProject
    For i = 1 To .VBComponents.count
      If .VBComponents(i).Type = numType Then
        count = count + 1
      End If
    Next i
  End With

  countComponents = count
End Function



Private Function absPath(ByVal pathFile As String, Optional ByVal curPath As String = "") As String
  ' �t�@�C���p�X���΃p�X�ɕϊ�
  Dim nameOS As String
  nameOS = application.OperatingSystem

  If curPath = "" Then
    curPath = ThisWorkbook.Path
  End If

  'replace Win backslash(Chr(92))
  pathFile = replace(pathFile, Chr(92), application.PathSeparator)
  
  'replace Mac ":"Chr(58)
  pathFile = replace(pathFile, ":", application.PathSeparator)
  
  'replace Unix "/"Chr(47)
  pathFile = replace(pathFile, "/", application.PathSeparator)

  ' �t�@�C���̖�������؂蕶���̏ꍇ�A��؂蕶�����폜
  If Right(pathFile, 1) = application.PathSeparator Then
    pathFile = Left(pathFile, Len(pathFile) - 1)
  End If

  Select Case Left(pathFile, 1)
  
    'Case1. . �Ŏn�܂�ꍇ(���Ύw��)
    Case ".":
  
      Select Case Left(pathFile, 2)
        
        ' Case1-1. ���Ύw�� "../" �Ή�
        Case "..":
            '../�̌���CurPaht�̃f�B���N�g�������
          Do While Left(pathFile, 2) = ".."
              curPath = Left(curPath, InStrRev(curPath, application.PathSeparator) - 1)
              pathFile = Right(pathFile, Len(pathFile) - 3) '../���폜
          Loop
    
        ' Case1-2. ���Ύw�� "./" �Ή�
        Case Else:
          pathFile = Right(pathFile, Len(pathFile) - 2) './���폜

      End Select

      absPath = curPath & application.PathSeparator & pathFile
      Exit Function

    
    'Case2. ��؂蕶���Ŏn�܂�ꍇ (��Ύw��)
    Case application.PathSeparator:
    
      ' Case2-1. Windows Network Drive ( chr(92) & chr(92) & "hoge")
      If Left(pathFile, 2) = Chr(92) & Chr(92) Then
        absPath = pathFile
        Exit Function
        
      ' (Windows only) Windows���΃p�X(\hoge)
      ElseIf Left(pathFile, 1) = Chr(92) Then
        absPath = curPath & pathFile
        Exit Function
        
      Else
      ' Case2-2. Mac/UNIX Absolute path (/hoge)
        absPath = pathFile
        Exit Function
      
      End If
    
  End Select


  'Case3. [A-z][0-9]�Ŏn�܂�ꍇ (Mac��Office�Ő��K�\�����g����� select���ɓ����ׂ�...)

  ' Case3-1.�h���C�u���^�[�Ή�("c:" & chr(92) �� "c" & chr(92) & chr(92)�ɂȂ��Ă��܂��̂ŏ����߂�)
  If nameOS Like "Windows *" And Left(pathFile, 2) Like "[A-z]" & application.PathSeparator Then
    absPath = replace(pathFile, application.PathSeparator, ":", 1, 1)
    Exit Function
  End If
 
  ' Case3-2. ���w�� "filename"�Ή�
  If Left(pathFile, 1) Like "[0-9]" Or Left(pathFile, 1) Like "[A-z]" Then
    absPath = curPath & application.PathSeparator & pathFile
    Exit Function
  Else
    MsgBox "Error[AbsPath]: fail to get absolute path."
  
  End If

End Function




Private Function checkExistFile(ByVal pathFile As String) As Boolean

  On Error GoTo Err_dir
  If dir(pathFile) = "" Then
    checkExistFile = False
  Else
    checkExistFile = True
  End If

  Exit Function

Err_dir:
  checkExistFile = False

End Function



'���X�g�t�@�C����z��ŕԂ�(�s����'(�R�����g)�̍s & ��s�͖�������)
Private Function list2array(ByVal pathFile As String) As Variant
    
  Dim nameOS As String
  nameOS = application.OperatingSystem
        
  '1. ���X�g�t�@�C���̓ǂݎ��
  Dim fp As Integer
  fp = FreeFile
  Open pathFile For Input As #fp
  
  '2. ���X�g�̔z��
  Dim arrayOutput() As String
  Dim countLine As Integer
  countLine = 0
  ReDim Preserve arrayOutput(countLine) ' �z��0�ŕԂ��ꍇ�����邽��
  
  Do Until EOF(fp)
    '���C�u�������X�g��1�s������
    Dim strLine As String
    Line Input #fp, strLine

    Dim isLf As Long
    isLf = InStr(strLine, vbLf)
    
    If nameOS Like "Windows *" And Not isLf = 0 Then
      'OS��Windows ���� ���X�g�� LF���܂܂��ꍇ (�t�@�C����UNIX�`��)
      '�t�@�C���S�̂�1�s�Ɍ����Ă��܂��B
      
      Dim arrayLineLF As Variant
      arrayLineLF = Split(strLine, vbLf)
    
      Dim i As Integer
      For i = 0 To UBound(arrayLineLF) - 1
        '�s���� '(�R�����g) �ł͂Ȃ� & ��s�ł͂Ȃ��ꍇ
        If Not Left(arrayLineLF(i), 1) = "'" And Len(arrayLineLF(i)) > 0 Then
      
          '�z��ւ̒ǉ�
          countLine = countLine + 1
          ReDim Preserve arrayOutput(countLine)
          arrayOutput(countLine - 1) = arrayLineLF(i)
        End If
      Next i
              
    
    Else
      'OS��Windows and �t�@�C����Windows�`�� (�ϊ��s�v)
      'OS��MacOS X and �t�@�C����UNIX�`�� (�ϊ��s�v)
      
      'OS��MacOS X and �t�@�C����Windows�`��
      ' vbCr�����W���[���t�@�C�����𔭌��ł��Ȃ��Ȃ�B
      strLine = replace(strLine, vbCr, "")
    
  
      '�s���� '(�R�����g) �ł͂Ȃ� & ��s�ł͂Ȃ��ꍇ
      If Not Left(strLine, 1) = "'" And Len(strLine) > 0 Then
      
        '�z��ւ̒ǉ�
        countLine = countLine + 1
        ReDim Preserve arrayOutput(countLine)
        arrayOutput(countLine - 1) = strLine
      End If
    
    End If
  Loop

  '3. ���X�g�t�@�C�������
  Close #fp
  
  '4. �߂�l��z��ŕԂ�
  list2array = arrayOutput
End Function

' loadMolude��2~3�����W���[����
Private Function getModuleList(ByRef pathConf, ByRef arrayModules, ByRef curPath As String)
  ' 2.0. ���W���[�����X�g�t�@�C���܂ł̑��΃p�X���΃p�X�Ƃ��Ď擾
  If FILEPATH_LIBLIST = "" Then
    curPath = ThisWorkbook.Path
  Else
    curPath = absPath(FILEPATH_LIBLIST)
  End If
  
  ' 2.1. ���W���[�����X�g�t�@�C���̐�΃p�X���擾
  pathConf = absPath(pathConf, curPath)
  
  ' 2.2. ���݃`�F�b�N
  Dim isExistList As Boolean
  isExistList = checkExistFile(pathConf)
  
  If isExistList = False Then
    getModuleList = "Error: ���C�u�������X�g" & pathConf & "�����݂��܂���B"
    Exit Function
  End If

  '3. ���W���[�����X�g�t�@�C���̓ǂݍ���&�z��
  arrayModules = list2array(pathConf)
  
  If UBound(arrayModules) = 0 Then
    getModuleList = "Error: ���C�u�������X�g�ɗL���ȃ��W���[���̋L�q�����݂��܂���B"
    Exit Function
  End If
End Function

' �V���[�g�J�b�g�̐ݒ� (Mac�ł� Macro�w��ł��Ȃ����ۂ�)
Private Sub setShortKey()
  If application.OperatingSystem Like "Windows *" Then
    application.MacroOptions Macro:="reloadModule", ShortcutKey:=SHORTKEY_RELOAD
    application.MacroOptions Macro:="exportModule", ShortcutKey:=SHORTKEY_EXPORT
  
  Else
    ' Mac OS X�̏ꍇ�̒���: ThisWorkbook.reloadModule�֐������}�N���t�@�C���𕡐��J���Ă���ƁA
    ' �Ō�ɊJ�����}�N���t�@�C���� ThisWorkbook.reloadModule�֐����Ăяo�����͗l�B
    ' (���̏ꍇ�A�}�N���ꗗ����'�Y���}�N���t�@�C��!reloadModule' ���Ăяo���Ă��������B)
    application.OnKey "^" & SHORTKEY_RELOAD, "reloadModule"
    application.OnKey "^" & SHORTKEY_EXPORT, "exportModule"
  End If
  
End Sub


'�V���[�g�J�b�g�ݒ�̍폜 (Mac�ł� Macro�w��ł��Ȃ����ۂ�)
Private Sub clearShortKey()
  If application.OperatingSystem Like "Windows *" Then
    application.MacroOptions Macro:="reloadModule", ShortcutKey:=""
    application.MacroOptions Macro:="exportModule", ShortcutKey:=""
  
  Else
    ' Mac OS X�̏ꍇ�̒���: ThisWorkbook.reloadModule�֐������}�N���t�@�C���𕡐��J���Ă���ƁA
    ' �Ō�ɊJ�����}�N���t�@�C���� ThisWorkbook.reloadModule�֐����N���A�����\���������Ǝv����(������)�B
    application.OnKey SHORTKEY_RELOAD, ""
    application.OnKey SHORTKEY_EXPORT, ""
  End If
  
End Sub

' ----- �����܂Ń��W���[���̃C���|�[�g�A�G�N�X�|�[�g�Ɋւ���@�\------
' ----- �������璷�����ǉ������A�悭�g���֐� -----


' �����Ȃ邨�܂��Ȃ�
Property Let Focus(ByVal Flag As Boolean)
    With application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Property

' �G���[����
' TODO ���߂ĉ��s�ڂŃG���[�����킩��悤�ɂ������B���̂܂܂���On Error GoTo 0�̂ق����G���[�ӏ��Ńf�o�b�O�ł��邩��g�����肪����
Public Sub ErrHandler()
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number & vbcrlf & "Detail: " & Err.Description
    End If
    Err.Clear
End Sub

' �u�b�N���J��(2�d�ɊJ���΍�)
Public Function bookOpen(ByVal Book As String, Optional ThisBookActivate As Boolean = True) As Boolean
  
  ' �u�b�N�����݂��Ȃ���ΏI��
  If Not IsExistFileDir(Book) Then
    MsgBox "�t�@�C�������݂��܂���B" & vbcrlf & "Path: " & Book
    ' �J���Ȃ��̂�False��Ԃ��B
    bookOpen = False
    Exit Function
  End If
  
    ' �����̓t���p�X�Ȃ̂Ńt�@�C�����݂̂𔲂��o��
    Dim bookName As String
    bookName = Mid(Book, InStrRev(Book, "\") + 1)
    
  If IsBookOpened(bookName) Then
      application.Workbooks(bookName).Activate
  Else
      Workbooks.Open Book
  End If
  
  ' �J�����u�b�N���A�N�e�B�u�ɂ��邩
  If ThisBookActivate Then ThisWorkbook.Activate

  ' �J������True��Ԃ�
  bookOpen = True
End Function

' �u�b�N�̑��݃`�F�b�N
Public Function IsExistFileDir(ByVal bookPath As String) As Boolean

  Dim a: a = dir(bookPath)
  If a <> "" Then
      IsExistFileDir = True
  Else
      IsExistFileDir = False
  End If
End Function

' �u�b�N�����łɊJ���Ă��邩�`�F�b�N
' (�Q�lURL https://goo.gl/EP8v4f)
Public Function IsBookOpened(ByVal bookName As String) As Boolean
    On Error Resume Next

    '// �ۑ��ς݂̃u�b�N������
    Open bookName For Append As #1
    Close #1

    ' �J���Ă���u�b�N������ƃG���[�ԍ�70���Ԃ�
    If Err.Number = 70 Then
        IsBookOpened = True ' �J����Ă���ꍇ
    Else
        IsBookOpened = False ' �J����Ă��Ȃ��ꍇ
    End If
    
    On Error GoTo 0
End Function

' �V�[�g�����݂��Ă��邩�`�F�b�N
' (�Q�lURL https://goo.gl/bg27SQ)
Public Function IsSheetExist(ByVal sheetName As String, Optional ByRef wb As Workbook) As Boolean
    Dim s As Excel.Worksheet

    On Error Resume Next
    If wb Is Nothing Then
        Set s = ActiveWorkbook.Sheets(sheetName)
    Else
        Set s = wb.Sheets(sheetName)
    End If
    On Error GoTo 0

    IsSheetExist = Not s Is Nothing
End Function

