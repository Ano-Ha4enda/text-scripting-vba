Attribute VB_Name = "main"
'@Folder("sampleMacros")
Option Explicit

Sub sub_1()
    Focus = True ' �}�N���������Ȃ邨�܂��Ȃ�
    On Error GoTo ErrHandler '�G���[�n���h��
    
    Dim num As Long
    num = 5000

    MsgBox func1(num)

ExitSub:
    Focus = False ' ���܂��Ȃ�������
    Exit Sub
    
ErrHandler:
    Call ErrHandler
    GoTo ExitSub
End Sub
