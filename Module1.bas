Attribute VB_Name = "Module1"
'
Sub Make_Img2_Db_fmt()
'
'2008.9.21 MI
'2023.5.5 Ojirou
'
' 実行中はExcelを隠し、終了メッセージを出す
'
    Dim i As Long, j As Long
    Dim myFile As String
    Dim fNo As Integer
'
    '
    '    ウィンドウを最小化する
    '
    With ActiveWindow
        .WindowState = xlMinimized
    End With
    '
    '   定型フォーマットを出力する
    '
    Worksheets("定型フォーマット").Activate
    '    数値を文字化する
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
    '    出力する
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
    '    ウィンドウを最大サイズに戻す
    '
    With ActiveWindow
        .WindowState = xlMaximized
    End With
 End Sub

