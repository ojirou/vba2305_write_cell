Attribute VB_Name = "Module1"
'
Sub Make_Img2_Db_fmt()
'
'2008.9.21 MI
'2023.5.5 Ojirou
'
' ���s����Excel���B���A�I�����b�Z�[�W���o��
'
    Dim i As Long, j As Long
    Dim myFile As String
    Dim fNo As Integer
'
    '
    '    �E�B���h�E���ŏ�������
    '
    With ActiveWindow
        .WindowState = xlMinimized
    End With
    '
    '   ��^�t�H�[�}�b�g���o�͂���
    '
    Worksheets("��^�t�H�[�}�b�g").Activate
    '    ���l�𕶎�������
    '
    For i = 3 To 99999
    For j = 2 To 4
    Select Case Cells(i, j)
        Case Is <= 9999999999#
            Cells(i, j) = "'" & Cells(i, j)
    End Select
    Next j
    If Cells(i + 1, 1) = "" Then Exit For
    Next i
    '
    '    �o�͂���
    '
    For i = 3 To 99999
        myFile = Cells(i, 2)
        fNo = FreeFile
        Open myFile For Output As #fNo
        Print #fNo, Cells(i, 3)
        Close #fNo
        If Cells(i + 1, 2) = "" Then Exit For
    Next i
'
    '
    '    �E�B���h�E���ő�T�C�Y�ɖ߂�
    '
    With ActiveWindow
        .WindowState = xlMaximized
    End With
 End Sub

