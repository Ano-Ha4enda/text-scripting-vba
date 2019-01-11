Attribute VB_Name = "main"
'@Folder("sampleMacros")
Option Explicit

Public Sub sub1()
    Focus = True ' マクロが早くなるおまじない
    On Error GoTo ErrHandler 'エラーハンドラ
    
    
    Focus = False ' おまじないを解除
End Sub
