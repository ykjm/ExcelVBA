Attribute VB_Name = "Module1"
Sub �����Z���̃f�[�^����()
    �s = ActiveCell.Row
    �� = ActiveCell.Column
   
    ���� = Cells(�s, ��)
    �Z���� = Selection.Cells.Count  '�I�����ꂽ�Z���͈͂̃Z�������擾
    If �Z���� < 2 Then
        MsgBox ("�I�����ꂽ�͈͂������Z���ł͂Ȃ��̂ŁA�����𒆎~���܂��B")
        Exit Sub
    End If
   
    For i = 1 To �Z���� - 1
        If Cells(�s + i, ��) <> "" Then
            ���� = ���� & vbCrLf & Cells(�s + i, ��)
        End If
    Next
    Cells(�s, ��) = ����
   
    '�s��l��(Active�Z���ȊO�̑I���Z�����폜)
    Range(Cells(�s + 1, ��), Cells(�s + �Z���� - 1, ��)).Delete Shift:=xlShiftUp
   
    Cells(�s, ��).Select
End Sub
