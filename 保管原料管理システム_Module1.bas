Attribute VB_Name = "Module1"

Option Explicit

'-----定数の指定-----
Public Enum R                       'R = 行
    出力f = 3: 出力l = 200          '出力の最初と最後
    棚6f = 3:   棚6l = 7            '棚6階の最初と最後
    棚5f = 10: 棚5l = 14            '棚5階の最初と最後
    dbf = 2                         'データベースの最初
End Enum
Public Enum C                       'C = 列
    出力f = 4: 出力l = 10           '出力の最初と最後
    棚6f = 13: 棚6l = 28            '棚6階の最初と最後
    棚5f = 13: 棚5l = 28            '棚5階の最初と最後
    dbf = 1: dbl = 7                'データベースの最初と最後
    dbGb = 5: dbGn = 6: dbHn = 7    'データベースの原料番号列,
                                    '原料名列,表示名称列
    seaTb = 6                       'searchシートの棚番号列
End Enum
Public Enum 棚len                   '棚len = 棚文字列内の記載位置
    階 = 1: 横 = 4: 縦 = 7
End Enum

Sub 検索(ByVal seaWord As String, ByVal dbCol As Long)
    
    Dim wsS As Worksheet    'searchシートの略名
    Dim wsD As Worksheet    'lab_zaikoシートの略名
    Dim Rdbl As Long        'データベースの最終行
    Dim dbRng As Range      '検索でヒットしたレコードのRange
    Dim seaRng As Range     'レコードを出力するRange
    Dim tgtWord As String   'データベースの検索対象
    Dim hitCnt As Long      '検索ヒット件数
    Dim i As Long           '繰り返し用
    
    '-----初期化-----
    Set wsS = Worksheets("search")
    Set wsD = Worksheets("lab_zaiko")
    Rdbl = wsD.ListObjects("labdb").ListRows.Count
    hitCnt = 0
    Call 検索クリア

    For i = R.dbf To Rdbl + 1
    
        '-----文字を大文字,カタカナ,半角に統一-----
        tgtWord = wsD.Cells(i, dbCol).Value
        tgtWord = StrConv(UCase(tgtWord), vbKatakana)
        tgtWord = StrConv(tgtWord, vbNarrow)
        seaWord = StrConv(UCase(seaWord), vbKatakana)
        seaWord = StrConv(seaWord, vbNarrow)
        
        '-----コピー元、貼り付け先のRangeの略名-----
        Set dbRng = wsD.Range(wsD.Cells(i, C.dbf), wsD.Cells(i, C.dbl))
        Set seaRng = wsS.Range(wsS.Cells(R.出力f + hitCnt, C.出力f), _
                     wsS.Cells(R.出力f + hitCnt, C.出力l))
        
        '-----検索と貼り付け-----
        If tgtWord Like "*" & seaWord & "*" Then
            seaRng.Value = dbRng.Value
            hitCnt = hitCnt + 1
        End If
    Next i

    '-----検索結果の件数とメッセージを表示-----
    If hitCnt = 0 Then
       MsgBox "抽出条件に一致するデータが存在しません", vbExclamation, "検索NG"
    Else
       MsgBox hitCnt & "件のデータを抽出しました", vbInformation, "検索OK"
    End If
    
    '-----棚に色付け-----
    Call 色付け(hitCnt)
    
End Sub

Sub 色付け(num As Long)
    Dim i As Long           '繰り返し用
    Dim tanaStr As String   '保管場所文字列
    Dim clrR As Long        '色付けするセルの行
    Dim clrC As Long        '色付けするセルの列
    Dim wsS As Worksheet    'searchシートの略名
    
    Set wsS = Worksheets("search")

    For i = 0 To num - 1
        
        '-----色付け行の取得-----
        tanaStr = wsS.Cells(R.出力f + i, C.seaTb).Value
        
        If Left(tanaStr, 棚len.階) = 6 Then                       '6階の6
            clrR = R.棚6f + Val(Mid(tanaStr, 棚len.縦, 2))        '2文字
        ElseIf Left(tanaStr, 棚len.階) = 5 Then                   '5階の5
            clrR = R.棚5f + Val(Mid(tanaStr, 棚len.縦, 2))        '2文字
        End If
        
        '-----色付け列の取得-----
        '-----偶数回に1度、行を空けたいので、1.5倍して小数を切り捨て-----
        clrC = C.棚5l - WorksheetFunction.RoundDown _
               (Val(Mid(tanaStr, 棚len.横, 2)) * 1.5, 0)          '2文字
        
        '-----枠内に収まっている場合に色付け-----
        If clrR > R.棚6f And clrR <= R.棚5l _
               And clrC >= C.棚5f And clrC < C.棚5l Then
            wsS.Cells(clrR, clrC).Interior.Color = RGB(255, 255, 0)
        End If
    
    Next i
End Sub

Sub 検索クリア()
    '-----検索結果のクリア-----
    Dim wsS As Worksheet
    Set wsS = Worksheets("search")

    With wsS
        .Range(.Cells(R.棚6f, C.棚6f), .Cells(R.棚5l, C.棚5l)) _
            .Interior.ColorIndex = 0
        
        With .Range(.Cells(R.出力f, C.出力f), .Cells(R.出力l, C.出力l))
            .Value = ""
            .ShrinkToFit = True
            .HorizontalAlignment = xlCenter
        End With
    End With
End Sub

Sub Open_Myform()
    '-----検索ボックスを開く-----
    検索ボックス.StartUpPosition = 0
    検索ボックス.Top = 350
    検索ボックス.Left = 350
    検索ボックス.Show
End Sub

