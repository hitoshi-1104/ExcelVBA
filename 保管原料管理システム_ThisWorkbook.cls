VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Workbook_open()

    '-----起動時にサーチシートの表示、検索結果のクリア-----
    With Worksheets("search")
        .Select
        .Cells(1, 1).Select
    End With
    Call 検索クリア

    '-----起動時にlab_zaikoシートのテーブル化-----
    Dim i As Long
    Dim tblrow As Long, tblcol As Long
    Dim tblrange As Range
    Dim tbl As ListObject
    
    With Worksheets("lab_zaiko")
        If .ListObjects.Count > 0 Then
            For i = .ListObjects.Count To 1 Step -1
                .ListObjects(i).TableStyle = ""
                .ListObjects(i).Unlist
            Next
        End If
        tblrow = .Cells(Rows.Count, 1).End(xlUp).Row
        tblcol = .Cells(1, Columns.Count).End(xlToLeft).Column
        Set tblrange = .Range(.Cells(1, 1), .Cells(tblrow, tblcol))
        .ListObjects.Add(, tblrange, , xlYes).Name = "labdb"
        .ListObjects("labdb").TableStyle = ""
    End With
    
End Sub
