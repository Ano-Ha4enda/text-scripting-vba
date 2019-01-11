Attribute VB_Name = "main"
'@Folder("sampleMacros")
Option Explicit

Sub sub_1()
    Focus = True ' マクロが早くなるおまじない
    On Error GoTo ErrHandler 'エラーハンドラ
    
    Dim num As Long
    num = 5000

    MsgBox func1(num)

ExitSub:
    Focus = False ' おまじないを解除
    Exit Sub
    
ErrHandler:
    Call ErrHandler
    GoTo ExitSub
End Sub
