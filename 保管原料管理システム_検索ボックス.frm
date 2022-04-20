VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 検索ボックス 
   Caption         =   "検索ボックス"
   ClientHeight    =   1800
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3936
   OleObjectBlob   =   "原料管理システム_検索ボックス.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "検索ボックス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim tb As String    'テキストボックス内容
Dim dbnum As Long   'DBの検索行

Private Sub UserForm_Initialize()   'オプションボタンの初期化
    OptionButton1 = True
End Sub
Private Sub OptionButton1_Click()   '原料番号
    dbnum = Val(C.dbGb)
    TextBox1.SetFocus
End Sub
Private Sub OptionButton2_Click()   '原料名
    dbnum = Val(C.dbGn)
    TextBox1.SetFocus
End Sub
Private Sub OptionButton3_Click()   '表示名称
    dbnum = Val(C.dbHn)
    TextBox1.SetFocus
End Sub

Private Sub Commandbutton1_Click()   '検索
    tb = TextBox1
    
    '-----未入力エラー回避-----
    If tb = "" Then
        MsgBox "検索文字を入力して下さい", vbExclamation, "未入力"
        TextBox1.SetFocus
        Exit Sub
    End If
    
    '-----テキストとオプションのデータを渡して検索の呼び出し-----
    Call 検索(tb, dbnum)
    TextBox1.SetFocus
    
End Sub

Private Sub Commandbutton2_Click()   '閉じる
    Unload Me
End Sub






