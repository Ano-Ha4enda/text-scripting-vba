Attribute VB_Name = "main"
'@Folder("sampleMacros")
Option Explicit

Public Sub sub1()
    Focus = True ' �}�N���������Ȃ邨�܂��Ȃ�
    On Error GoTo ErrHandler '�G���[�n���h��
    
    
    Focus = False ' ���܂��Ȃ�������
End Sub
