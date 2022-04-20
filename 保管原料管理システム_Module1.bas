Attribute VB_Name = "Module1"

Option Explicit

'-----�萔�̎w��-----
Public Enum R                       'R = �s
    �o��f = 3: �o��l = 200          '�o�͂̍ŏ��ƍŌ�
    �I6f = 3:   �I6l = 7            '�I6�K�̍ŏ��ƍŌ�
    �I5f = 10: �I5l = 14            '�I5�K�̍ŏ��ƍŌ�
    dbf = 2                         '�f�[�^�x�[�X�̍ŏ�
End Enum
Public Enum C                       'C = ��
    �o��f = 4: �o��l = 10           '�o�͂̍ŏ��ƍŌ�
    �I6f = 13: �I6l = 28            '�I6�K�̍ŏ��ƍŌ�
    �I5f = 13: �I5l = 28            '�I5�K�̍ŏ��ƍŌ�
    dbf = 1: dbl = 7                '�f�[�^�x�[�X�̍ŏ��ƍŌ�
    dbGb = 5: dbGn = 6: dbHn = 7    '�f�[�^�x�[�X�̌����ԍ���,
                                    '��������,�\�����̗�
    seaTb = 6                       'search�V�[�g�̒I�ԍ���
End Enum
Public Enum �Ilen                   '�Ilen = �I��������̋L�ڈʒu
    �K = 1: �� = 4: �c = 7
End Enum

Sub ����(ByVal seaWord As String, ByVal dbCol As Long)
    
    Dim wsS As Worksheet    'search�V�[�g�̗���
    Dim wsD As Worksheet    'lab_zaiko�V�[�g�̗���
    Dim Rdbl As Long        '�f�[�^�x�[�X�̍ŏI�s
    Dim dbRng As Range      '�����Ńq�b�g�������R�[�h��Range
    Dim seaRng As Range     '���R�[�h���o�͂���Range
    Dim tgtWord As String   '�f�[�^�x�[�X�̌����Ώ�
    Dim hitCnt As Long      '�����q�b�g����
    Dim i As Long           '�J��Ԃ��p
    
    '-----������-----
    Set wsS = Worksheets("search")
    Set wsD = Worksheets("lab_zaiko")
    Rdbl = wsD.ListObjects("labdb").ListRows.Count
    hitCnt = 0
    Call �����N���A

    For i = R.dbf To Rdbl + 1
    
        '-----������啶��,�J�^�J�i,���p�ɓ���-----
        tgtWord = wsD.Cells(i, dbCol).Value
        tgtWord = StrConv(UCase(tgtWord), vbKatakana)
        tgtWord = StrConv(tgtWord, vbNarrow)
        seaWord = StrConv(UCase(seaWord), vbKatakana)
        seaWord = StrConv(seaWord, vbNarrow)
        
        '-----�R�s�[���A�\��t�����Range�̗���-----
        Set dbRng = wsD.Range(wsD.Cells(i, C.dbf), wsD.Cells(i, C.dbl))
        Set seaRng = wsS.Range(wsS.Cells(R.�o��f + hitCnt, C.�o��f), _
                     wsS.Cells(R.�o��f + hitCnt, C.�o��l))
        
        '-----�����Ɠ\��t��-----
        If tgtWord Like "*" & seaWord & "*" Then
            seaRng.Value = dbRng.Value
            hitCnt = hitCnt + 1
        End If
    Next i

    '-----�������ʂ̌����ƃ��b�Z�[�W��\��-----
    If hitCnt = 0 Then
       MsgBox "���o�����Ɉ�v����f�[�^�����݂��܂���", vbExclamation, "����NG"
    Else
       MsgBox hitCnt & "���̃f�[�^�𒊏o���܂���", vbInformation, "����OK"
    End If
    
    '-----�I�ɐF�t��-----
    Call �F�t��(hitCnt)
    
End Sub

Sub �F�t��(num As Long)
    Dim i As Long           '�J��Ԃ��p
    Dim tanaStr As String   '�ۊǏꏊ������
    Dim clrR As Long        '�F�t������Z���̍s
    Dim clrC As Long        '�F�t������Z���̗�
    Dim wsS As Worksheet    'search�V�[�g�̗���
    
    Set wsS = Worksheets("search")

    For i = 0 To num - 1
        
        '-----�F�t���s�̎擾-----
        tanaStr = wsS.Cells(R.�o��f + i, C.seaTb).Value
        
        If Left(tanaStr, �Ilen.�K) = 6 Then                       '6�K��6
            clrR = R.�I6f + Val(Mid(tanaStr, �Ilen.�c, 2))        '2����
        ElseIf Left(tanaStr, �Ilen.�K) = 5 Then                   '5�K��5
            clrR = R.�I5f + Val(Mid(tanaStr, �Ilen.�c, 2))        '2����
        End If
        
        '-----�F�t����̎擾-----
        '-----�������1�x�A�s���󂯂����̂ŁA1.5�{���ď�����؂�̂�-----
        clrC = C.�I5l - WorksheetFunction.RoundDown _
               (Val(Mid(tanaStr, �Ilen.��, 2)) * 1.5, 0)          '2����
        
        '-----�g���Ɏ��܂��Ă���ꍇ�ɐF�t��-----
        If clrR > R.�I6f And clrR <= R.�I5l _
               And clrC >= C.�I5f And clrC < C.�I5l Then
            wsS.Cells(clrR, clrC).Interior.Color = RGB(255, 255, 0)
        End If
    
    Next i
End Sub

Sub �����N���A()
    '-----�������ʂ̃N���A-----
    Dim wsS As Worksheet
    Set wsS = Worksheets("search")

    With wsS
        .Range(.Cells(R.�I6f, C.�I6f), .Cells(R.�I5l, C.�I5l)) _
            .Interior.ColorIndex = 0
        
        With .Range(.Cells(R.�o��f, C.�o��f), .Cells(R.�o��l, C.�o��l))
            .Value = ""
            .ShrinkToFit = True
            .HorizontalAlignment = xlCenter
        End With
    End With
End Sub

Sub Open_Myform()
    '-----�����{�b�N�X���J��-----
    �����{�b�N�X.StartUpPosition = 0
    �����{�b�N�X.Top = 350
    �����{�b�N�X.Left = 350
    �����{�b�N�X.Show
End Sub

