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

'ライブラリリストの設定 (設置フォルダはワークブックと同じディレクトリ)
Const FILENAME_LIBLIST As String = "libdef.txt" 'ライブラリリストのファイル名
Const FILEPATH_LIBLIST As String = "" 'エクセルファイルから見たライブラリリストの相対パス(頭にパス区切り文字は入れないこと)

'ワークブック オープン時に実行する(True) / しない(False)
Const ENABLE_WORKBOOK_OPEN As Boolean = False

'ショートカットキー
Const SHORTKEY_RELOAD As String = "r" 'ctrl + r
Const SHORTKEY_EXPORT As String = "e" 'ctrl + e

'----------------------------- Workbook_open() ---------------

'ワークブック オープン時に実行
Private Sub Workbook_Open()
  Call setShortKey
  If ENABLE_WORKBOOK_OPEN = True Then
    Call reloadModule
  End If
 End Sub

'ワークブック クローズ時に実行
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Call clearShortKey
 End Sub



'----------------------------- public Subs/Functions ---------------

Public Sub reloadModule()
Attribute reloadModule.VB_ProcData.VB_Invoke_Func = "r\n14"
  '手動リロード用 Public関数
  
  Dim msgError As String
  msgError = loadModule("." & application.PathSeparator & FILENAME_LIBLIST)
  
  If Len(msgError) > 0 Then
    MsgBox msgError
  End If
End Sub


Public Sub exportModules()
  '手動export用 Public関数
  Dim arrayModules As Variant
  Dim i As Integer
  Dim message As String
  Dim msgError As String
  Dim curPath As String

  'モジュールリストファイルの存在確認、読み込み、配列化
  msgError = getModuleList(FILENAME_LIBLIST, arrayModules, curPath)
  If Len(msgError) > 0 Then GoTo ErrHandler
  'モジュールリスト
  For i = 0 To UBound(arrayModules) - 1
    arrayModules(i) = absPath(arrayModules(i), curPath)
  Next
  
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents

    'エクスポートするファイルのフルパス
    Dim pathModule As String
    pathModule = "" '初期化

    ' 拡張子を指定
    Dim moduleType As String
    If component.Type = 1 Then moduleType = ".bas"
    If component.Type = 2 Then moduleType = ".cls"

    ' ライブラリのフルパスの中に、モジュールと同じ名前のファイルがあるかチェック
    For i = 0 To UBound(arrayModules) - 1
      If InStr(arrayModules(i), application.PathSeparator & component.Name & moduleType) > 0 Then
        pathModule = arrayModules(i)
        Exit For
      End If
    Next i

    ' TODO モジュールが無ければエラー。シートオブジェクトとモジュールを切り分ける方法が分かれば実装してほしい
    ' モジュールをエクスポート
    If pathModule <> "" Then exportModule component, pathModule, message
      
    Next

  ' 成功した場合、エクスポートファイル一覧を表示
  MsgBox message
  Exit Sub

ErrHandler:
  MsgBox msgError
End Sub




'----------------------------- main Subs/Functions ---------------

Private Function loadModule(ByVal pathConf As String) As String
  'Main: モジュールリストファイルに書いてある外部ライブラリを読み込む。

  '1. 全モジュールを削除
  Dim isClear As Boolean
  isClear = clearModules
  
  If isClear = False Then
    loadModule = "Error: 標準モジュールの全削除に失敗しました。"
    Exit Function
  End If
  
  
  '2. モジュールリストファイルの存在確認
  '3. モジュールリストファイルの読み込み&配列化
  Dim arrayModules As Variant
  Dim msgError As String
  Dim curPath As String
  msgError = getModuleList(pathConf, arrayModules, curPath)
  If msgError <> "" Then GoTo msgErr
  
  '4. 各モジュールファイル読み込み
  Dim i As Integer
  
  ' 配列は0始まり。(最大値: 配列個数-1)
  For i = 0 To UBound(arrayModules) - 1
    Dim pathModule As String
    pathModule = arrayModules(i)
    
    '4.1. モジュールリストファイルの存在確認
    ' 4.1.1. モジュールリストファイルの絶対パスを取得
    pathModule = absPath(pathModule, curPath)
  
    ' 4.1.2. 存在チェック
    Dim isExistModule As Boolean
    isExistModule = checkExistFile(pathModule)
  
    '4.2. モジュール読み込み
    If isExistModule = True Then
      ThisWorkbook.VBProject.VBComponents.Import pathModule
    Else
      msgError = msgError & pathModule & " は存在しません。" & vbcrlf
    End If
  Next i
  
msgErr:
  loadModule = msgError

End Function



'----------------------------- Functions / Subs ---------------

Private Sub exportModule(ByVal component As Object, ByVal pathModule As String, ByRef message As String)
  
  component.Export pathModule
  message = message & component.Name & " を " & pathModule & " として保存しました。" & vbcrlf
  
End Sub




'----------------------------- common Functions / Subs ---------------
Private Function clearModules() As Boolean
  '標準モジュール/クラスモジュール初期化(全削除)
  
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents

    ' このモジュール自身は削除しない
    If component.Name = "moduleIO" Then GoTo Continue

    '標準モジュール(Type=1) / クラスモジュール(Type=2)を全て削除
    If component.Type = 1 Or component.Type = 2 Then
      ThisWorkbook.VBProject.VBComponents.Remove component
    End If
    
Continue:
  Next component
  
  '標準モジュール/クラスモジュールの合計数が1(このモジュール自身のみ)であればOK
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
  '存在する標準モジュール/クラスモジュールの数を数える
  
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
  ' ファイルパスを絶対パスに変換
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

  ' ファイルの末尾が区切り文字の場合、区切り文字を削除
  If Right(pathFile, 1) = application.PathSeparator Then
    pathFile = Left(pathFile, Len(pathFile) - 1)
  End If

  Select Case Left(pathFile, 1)
  
    'Case1. . で始まる場合(相対指定)
    Case ".":
  
      Select Case Left(pathFile, 2)
        
        ' Case1-1. 相対指定 "../" 対応
        Case "..":
            '../の個数分CurPahtのディレクトリを削る
          Do While Left(pathFile, 2) = ".."
              curPath = Left(curPath, InStrRev(curPath, application.PathSeparator) - 1)
              pathFile = Right(pathFile, Len(pathFile) - 3) '../を削除
          Loop
    
        ' Case1-2. 相対指定 "./" 対応
        Case Else:
          pathFile = Right(pathFile, Len(pathFile) - 2) './を削除

      End Select

      absPath = curPath & application.PathSeparator & pathFile
      Exit Function

    
    'Case2. 区切り文字で始まる場合 (絶対指定)
    Case application.PathSeparator:
    
      ' Case2-1. Windows Network Drive ( chr(92) & chr(92) & "hoge")
      If Left(pathFile, 2) = Chr(92) & Chr(92) Then
        absPath = pathFile
        Exit Function
        
      ' (Windows only) Windows相対パス(\hoge)
      ElseIf Left(pathFile, 1) = Chr(92) Then
        absPath = curPath & pathFile
        Exit Function
        
      Else
      ' Case2-2. Mac/UNIX Absolute path (/hoge)
        absPath = pathFile
        Exit Function
      
      End If
    
  End Select


  'Case3. [A-z][0-9]で始まる場合 (Mac版Officeで正規表現が使えれば select文に入れるべき...)

  ' Case3-1.ドライブレター対応("c:" & chr(92) が "c" & chr(92) & chr(92)になってしまうので書き戻す)
  If nameOS Like "Windows *" And Left(pathFile, 2) Like "[A-z]" & application.PathSeparator Then
    absPath = replace(pathFile, application.PathSeparator, ":", 1, 1)
    Exit Function
  End If
 
  ' Case3-2. 無指定 "filename"対応
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



'リストファイルを配列で返す(行頭が'(コメント)の行 & 空行は無視する)
Private Function list2array(ByVal pathFile As String) As Variant
    
  Dim nameOS As String
  nameOS = application.OperatingSystem
        
  '1. リストファイルの読み取り
  Dim fp As Integer
  fp = FreeFile
  Open pathFile For Input As #fp
  
  '2. リストの配列化
  Dim arrayOutput() As String
  Dim countLine As Integer
  countLine = 0
  ReDim Preserve arrayOutput(countLine) ' 配列0で返す場合があるため
  
  Do Until EOF(fp)
    'ライブラリリストを1行ずつ処理
    Dim strLine As String
    Line Input #fp, strLine

    Dim isLf As Long
    isLf = InStr(strLine, vbLf)
    
    If nameOS Like "Windows *" And Not isLf = 0 Then
      'OSがWindows かつ リストに LFが含まれる場合 (ファイルがUNIX形式)
      'ファイル全体で1行に見えてしまう。
      
      Dim arrayLineLF As Variant
      arrayLineLF = Split(strLine, vbLf)
    
      Dim i As Integer
      For i = 0 To UBound(arrayLineLF) - 1
        '行頭が '(コメント) ではない & 空行ではない場合
        If Not Left(arrayLineLF(i), 1) = "'" And Len(arrayLineLF(i)) > 0 Then
      
          '配列への追加
          countLine = countLine + 1
          ReDim Preserve arrayOutput(countLine)
          arrayOutput(countLine - 1) = arrayLineLF(i)
        End If
      Next i
              
    
    Else
      'OSがWindows and ファイルがWindows形式 (変換不要)
      'OSがMacOS X and ファイルがUNIX形式 (変換不要)
      
      'OSがMacOS X and ファイルがWindows形式
      ' vbCrがモジュールファイル名を発見できなくなる。
      strLine = replace(strLine, vbCr, "")
    
  
      '行頭が '(コメント) ではない & 空行ではない場合
      If Not Left(strLine, 1) = "'" And Len(strLine) > 0 Then
      
        '配列への追加
        countLine = countLine + 1
        ReDim Preserve arrayOutput(countLine)
        arrayOutput(countLine - 1) = strLine
      End If
    
    End If
  Loop

  '3. リストファイルを閉じる
  Close #fp
  
  '4. 戻り値を配列で返す
  list2array = arrayOutput
End Function

' loadMoludeの2~3をモジュール化
Private Function getModuleList(ByRef pathConf, ByRef arrayModules, ByRef curPath As String)
  ' 2.0. モジュールリストファイルまでの相対パスを絶対パスとして取得
  If FILEPATH_LIBLIST = "" Then
    curPath = ThisWorkbook.Path
  Else
    curPath = absPath(FILEPATH_LIBLIST)
  End If
  
  ' 2.1. モジュールリストファイルの絶対パスを取得
  pathConf = absPath(pathConf, curPath)
  
  ' 2.2. 存在チェック
  Dim isExistList As Boolean
  isExistList = checkExistFile(pathConf)
  
  If isExistList = False Then
    getModuleList = "Error: ライブラリリスト" & pathConf & "が存在しません。"
    Exit Function
  End If

  '3. モジュールリストファイルの読み込み&配列化
  arrayModules = list2array(pathConf)
  
  If UBound(arrayModules) = 0 Then
    getModuleList = "Error: ライブラリリストに有効なモジュールの記述が存在しません。"
    Exit Function
  End If
End Function

' ショートカットの設定 (Macでは Macro指定できないっぽい)
Private Sub setShortKey()
  If application.OperatingSystem Like "Windows *" Then
    application.MacroOptions Macro:="reloadModule", ShortcutKey:=SHORTKEY_RELOAD
    application.MacroOptions Macro:="exportModule", ShortcutKey:=SHORTKEY_EXPORT
  
  Else
    ' Mac OS Xの場合の注意: ThisWorkbook.reloadModule関数を持つマクロファイルを複数開いていると、
    ' 最後に開いたマクロファイルの ThisWorkbook.reloadModule関数が呼び出される模様。
    ' (その場合、マクロ一覧から'該当マクロファイル!reloadModule' を呼び出してください。)
    application.OnKey "^" & SHORTKEY_RELOAD, "reloadModule"
    application.OnKey "^" & SHORTKEY_EXPORT, "exportModule"
  End If
  
End Sub


'ショートカット設定の削除 (Macでは Macro指定できないっぽい)
Private Sub clearShortKey()
  If application.OperatingSystem Like "Windows *" Then
    application.MacroOptions Macro:="reloadModule", ShortcutKey:=""
    application.MacroOptions Macro:="exportModule", ShortcutKey:=""
  
  Else
    ' Mac OS Xの場合の注意: ThisWorkbook.reloadModule関数を持つマクロファイルを複数開いていると、
    ' 最後に開いたマクロファイルの ThisWorkbook.reloadModule関数がクリアされる可能性が高いと思われる(未検証)。
    application.OnKey SHORTKEY_RELOAD, ""
    application.OnKey SHORTKEY_EXPORT, ""
  End If
  
End Sub

' ----- ここまでモジュールのインポート、エクスポートに関する機能------
' ----- ここから長島が追加した、よく使う関数 -----


' 早くなるおまじない
Property Let Focus(ByVal Flag As Boolean)
    With application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Property

' エラー処理
' TODO せめて何行目でエラーかがわかるようにしたい。このままだとOn Error GoTo 0のほうがエラー箇所でデバッグできるから使い勝手がいい
Public Sub ErrHandler()
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number & vbcrlf & "Detail: " & Err.Description
    End If
    Err.Clear
End Sub

' ブックを開く(2重に開く対策)
Public Function bookOpen(ByVal Book As String, Optional ThisBookActivate As Boolean = True) As Boolean
  
  ' ブックが存在しなければ終了
  If Not IsExistFileDir(Book) Then
    MsgBox "ファイルが存在しません。" & vbcrlf & "Path: " & Book
    ' 開けないのでFalseを返す。
    bookOpen = False
    Exit Function
  End If
  
    ' 引数はフルパスなのでファイル名のみを抜き出す
    Dim bookName As String
    bookName = Mid(Book, InStrRev(Book, "\") + 1)
    
  If IsBookOpened(bookName) Then
      application.Workbooks(bookName).Activate
  Else
      Workbooks.Open Book
  End If
  
  ' 開いたブックをアクティブにするか
  If ThisBookActivate Then ThisWorkbook.Activate

  ' 開いたらTrueを返す
  bookOpen = True
End Function

' ブックの存在チェック
Public Function IsExistFileDir(ByVal bookPath As String) As Boolean

  Dim a: a = dir(bookPath)
  If a <> "" Then
      IsExistFileDir = True
  Else
      IsExistFileDir = False
  End If
End Function

' ブックをすでに開いているかチェック
' (参考URL https://goo.gl/EP8v4f)
Public Function IsBookOpened(ByVal bookName As String) As Boolean
    On Error Resume Next

    '// 保存済みのブックか判定
    Open bookName For Append As #1
    Close #1

    ' 開いているブックがあるとエラー番号70が返る
    If Err.Number = 70 Then
        IsBookOpened = True ' 開かれている場合
    Else
        IsBookOpened = False ' 開かれていない場合
    End If
    
    On Error GoTo 0
End Function

' シートが存在しているかチェック
' (参考URL https://goo.gl/bg27SQ)
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

