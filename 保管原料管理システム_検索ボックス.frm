VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �����{�b�N�X 
   Caption         =   "�����{�b�N�X"
   ClientHeight    =   1800
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3936
   OleObjectBlob   =   "�����Ǘ��V�X�e��_�����{�b�N�X.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�����{�b�N�X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim tb As String    '�e�L�X�g�{�b�N�X���e
Dim dbnum As Long   'DB�̌����s

Private Sub UserForm_Initialize()   '�I�v�V�����{�^���̏�����
    OptionButton1 = True
End Sub
Private Sub OptionButton1_Click()   '�����ԍ�
    dbnum = Val(C.dbGb)
    TextBox1.SetFocus
End Sub
Private Sub OptionButton2_Click()   '������
    dbnum = Val(C.dbGn)
    TextBox1.SetFocus
End Sub
Private Sub OptionButton3_Click()   '�\������
    dbnum = Val(C.dbHn)
    TextBox1.SetFocus
End Sub

Private Sub Commandbutton1_Click()   '����
    tb = TextBox1
    
    '-----�����̓G���[���-----
    If tb = "" Then
        MsgBox "������������͂��ĉ�����", vbExclamation, "������"
        TextBox1.SetFocus
        Exit Sub
    End If
    
    '-----�e�L�X�g�ƃI�v�V�����̃f�[�^��n���Č����̌Ăяo��-----
    Call ����(tb, dbnum)
    TextBox1.SetFocus
    
End Sub

Private Sub Commandbutton2_Click()   '����
    Unload Me
End Sub






