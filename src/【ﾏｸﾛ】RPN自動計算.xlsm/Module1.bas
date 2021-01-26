Attribute VB_Name = "Module1"
Option Explicit

Public Const StartRow As Long = 13
Public Const StartColumn As Long = 3
Public Const cbStartColumn As Long = 4
Public Const cbEndColumn As Long = 8
Public Const RPNColumn As Long = 9
Public Const RankRow As Long = 9

Sub Create_CheckBox()

    Application.ScreenUpdating = False

    Dim EndRow As Long
    Dim Result As String
    Dim fig As Shape
    Dim i As Long
    Dim j As Long
    Dim cbLeft As Long
    Dim cbTop As Long
    Dim cbWidth As Long
    Dim cbHeight As Long

    EndRow = Cells(Rows.Count, StartColumn).End(xlUp).row

    If EndRow < StartRow Then
        
        MsgBox "要素作業を入力してください。", vbExclamation
        Exit Sub
    
    End If

    If Application.WorksheetFunction.CountBlank(Range(Cells(StartRow, StartColumn), Cells(EndRow, StartColumn))) = 0 Then

        GoTo Skip

    End If

    If Range(Cells(StartRow, StartColumn), Cells(EndRow, StartColumn)).SpecialCells(xlCellTypeBlanks).Count <> 0 = True Then

        Result = MsgBox("空白行を詰めます。" & vbCrLf & "よろしいですか？", vbYesNo + vbExclamation)

        If Result = vbYes Then

            Range(Cells(StartRow, StartColumn), Cells(EndRow, StartColumn)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete xlShiftUp
            EndRow = Cells(Rows.Count, StartColumn).End(xlUp).row

        ElseIf Result = vbNo Then

            MsgBox "空白行を詰めてください。", vbExclamation
            Exit Sub

        End If

    End If

Skip:

      Range(Cells(StartRow, StartColumn), Cells(EndRow, StartColumn)).Interior.Color = RGB(255, 255, 204)

    With Cells

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        With .Font

            .Size = 14
            .Name = "Meiryo UI"

        End With

    End With

    Rows.RowHeight = 24
    Columns.AutoFit

    If Columns("C").ColumnWidth < 13 Then

        Columns("C").ColumnWidth = 13

    End If

    If Columns("F").ColumnWidth < 20 Then

        Columns("F").ColumnWidth = 20

    End If

    If Columns("I").ColumnWidth < 15 Then

        Columns("I").ColumnWidth = 15

    End If

    If Columns("J").ColumnWidth < 50 Then

        Columns("J").ColumnWidth = 50

    End If

    Range(Cells(StartRow, StartColumn).Offset(0, -1), Cells(EndRow, StartColumn).Offset(0, -1)).HorizontalAlignment = xlCenter
    Range(Cells(StartRow, StartColumn), Cells(EndRow, StartColumn)).HorizontalAlignment = xlLeft
    Range(Cells(StartRow, StartColumn).Offset(0, -1), Cells(EndRow, RPNColumn).Offset(0, 1)).Borders.LineStyle = xlContinuous
    Range(Cells(StartRow, StartColumn), Cells(EndRow, StartColumn)).Offset(0, -1).Borders(xlEdgeLeft).Weight = xlMedium

    With Range(Cells(StartRow, StartColumn), Cells(EndRow, StartColumn))

        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium

    End With

    With Range(Cells(StartRow, RPNColumn), Cells(EndRow, RPNColumn))

        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium

    End With

    Range(Cells(StartRow, RPNColumn), Cells(EndRow, RPNColumn)).Offset(0, -1).Borders(xlEdgeRight).Weight = xlMedium
    Range(Cells(EndRow, StartColumn).Offset(0, -1), Cells(EndRow, RPNColumn).Offset(0, 1)).Borders(xlEdgeBottom).Weight = xlMedium

    For i = StartRow To EndRow

        Cells(i, StartColumn).Offset(0, -1) = (i - StartRow) + 1

        For j = cbStartColumn To cbEndColumn

            With Cells(i, j)

                cbLeft = .Left + (.Width / 2) - 6
                cbTop = .Top
                cbWidth = (.Offset(0, 1).Left - .Left) / 4
                cbHeight = .Height
                ActiveSheet.CheckBoxes.Add(cbLeft, cbTop, cbWidth, cbHeight).Select

                With Selection

                    .Text = ""
                    .LinkedCell = Cells(i, j).Address
                    Cells(i, j) = "False"

                End With

            End With

        Next j

    Next i

    Range(Cells(StartRow, cbStartColumn), Cells(EndRow, cbEndColumn)).Font.ColorIndex = 2

    Application.ScreenUpdating = True

End Sub

Sub Caluculate_RPN()

    Application.ScreenUpdating = False

    Dim EndRow As Long
    Dim i As Long
    Dim j As Long
    Dim ans1 As Long
    Dim ans2 As Long
    Dim ans3 As Long
    Dim rank As String
    Dim fc As FormatCondition

    EndRow = Cells(Rows.Count, StartColumn).End(xlUp).row

    If EndRow < StartRow Then

        Exit Sub

    End If

    Range(Cells(StartRow, RPNColumn), Cells(EndRow, RPNColumn)).FormatConditions.Delete

    For i = StartRow To EndRow

        ans1 = 0
        ans2 = 0
        ans3 = 0

        For j = cbStartColumn To cbEndColumn

            rank = Cells(RankRow, j)

            If rank = "" Then

                MsgBox "評価点ランクが入力されているか確認してください。", vbExclamation
                Exit Sub

            End If

            If j = cbStartColumn Or j = cbStartColumn + 1 Then

                Call Judge_Check(ans1, i, j, rank)

            ElseIf j = (cbStartColumn + cbEndColumn) / 2 Then

                Call Judge_Check(ans2, i, j, rank)

            ElseIf j = cbEndColumn - 1 Or j = cbEndColumn Then

                Call Judge_Check(ans3, i, j, rank)

            End If

        Next j

        Cells(i, RPNColumn) = ans1 * ans2 * ans3
        Set fc = Cells(i, RPNColumn).FormatConditions.Add(xlCellValue, xlGreater, 16)
        fc.Interior.Color = RGB(255, 192, 0)

    Next i

    Application.ScreenUpdating = True

End Sub

Sub Judge_Check(ans, row, column, rank)

    Select Case rank

        Case "1～2"

            If Cells(row, column) = True Then

                ans = ans + 2

            ElseIf Cells(row, column) = False Then

                ans = ans + 1
            
            End If

        Case "2～3"

            If Cells(row, column) = True Then

                ans = ans + 3

            ElseIf Cells(row, column) = False Then

                ans = ans + 2
            
            End If
    
    End Select

End Sub

Sub Sheet_Clear()

    Application.ScreenUpdating = False

    Dim cb As CheckBox

    Range(Cells(StartRow, StartColumn).Offset(0, -1), Cells(Rows.Count, RPNColumn).Offset(0, 1)).Clear

    For Each cb In ActiveSheet.CheckBoxes

        cb.Delete

    Next cb

    Application.ScreenUpdating = True

End Sub
