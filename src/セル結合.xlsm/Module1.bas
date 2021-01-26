Attribute VB_Name = "Module1"
Sub 複数セルのデータ統合()
    行 = ActiveCell.Row
    列 = ActiveCell.Column
   
    文字 = Cells(行, 列)
    セル数 = Selection.Cells.Count  '選択されたセル範囲のセル数を取得
    If セル数 < 2 Then
        MsgBox ("選択された範囲が複数セルではないので、処理を中止します。")
        Exit Sub
    End If
   
    For i = 1 To セル数 - 1
        If Cells(行 + i, 列) <> "" Then
            文字 = 文字 & vbCrLf & Cells(行 + i, 列)
        End If
    Next
    Cells(行, 列) = 文字
   
    '行上詰め(Activeセル以外の選択セルを削除)
    Range(Cells(行 + 1, 列), Cells(行 + セル数 - 1, 列)).Delete Shift:=xlShiftUp
   
    Cells(行, 列).Select
End Sub
