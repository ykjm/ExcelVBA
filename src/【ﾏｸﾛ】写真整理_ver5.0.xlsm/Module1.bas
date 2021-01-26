Attribute VB_Name = "Module1"
Option Explicit

Sub InsertPhoto()

    Dim fName As Variant
    Dim i As Long
    Dim Pict As Picture
    Dim object As Long
    Dim ans As String
    Dim rng As Range

    Cells.Font.ColorIndex = 1
    fName = Application.GetOpenFilename(Filefilter:="�摜�t�@�C��, *.jpg; *.png; *.gif; *.tif;*.bmp", MultiSelect:=True)

    If IsArray(fName) <> True Then

        MsgBox "�L�����Z�����܂���"
        Exit Sub

    End If

    object = ActiveSheet.Pictures.Count + 1
    Cells(object, 1).Activate

    If IsArray(fName) Then

        Application.ScreenUpdating = False

        BubbleSort fName, True

            For i = 1 To UBound(fName)

                Set Pict = ActiveSheet.Pictures.Insert(fName(i))

                    With Pict

                        .TopLeftCell = ActiveCell
                        .ShapeRange.LockAspectRatio = msoTrue
                        '���L�����ꂩ���R�����g�A�E�g
                        .ShapeRange.Height = ActiveCell.Height '�Z���̍����Ƀ��T�C�Y
                        '.ShapeRange.Width = ActiveCell.Width '�Z���̕��Ƀ��T�C�Y

                        With ActiveCell.Offset(0, 1)

                            .Value = fName(i) '�t�@�C�����̏�������

                            If i = 1 Then

                                .Font.ColorIndex = 3

                            End If
                        
                        End With
                    
                    End With
                
                ActiveCell.Offset(1, 0).Activate
                Application.StatusBar = "�������F" & i & "/" & UBound(fName) & "����"
                
            Next i

    End If

    With Application

        .StatusBar = False
        .ScreenUpdating = True

    End With

    Set Pict = Nothing

    ActiveSheet.PageSetup.PrintArea = Range(Cells(1, 1), Cells(object + UBound(fName) - LBound(fName), 1)).Address
    MsgBox i - 1 & "���̉摜��}�����܂���", vbInformation

End Sub

'�l�̓���ւ�
Public Sub Swap(ByRef Dat1 As Variant, ByRef Dat2 As Variant)

    Dim varBuf As Variant
    varBuf = Dat1
    Dat1 = Dat2
    Dat2 = varBuf

End Sub

'�z��̃o�u���\�[�g
Public Sub BubbleSort(ByRef aryDat As Variant, Optional ByVal SortAsc As Boolean = True)

    Dim i As Long
    Dim j As Long

        For i = LBound(aryDat) To UBound(aryDat) - 1

            For j = LBound(aryDat) To LBound(aryDat) + UBound(aryDat) - i - 1

                If aryDat(IIf(SortAsc, j, j + 1)) > aryDat(IIf(SortAsc, j + 1, j)) Then

                    Call Swap(aryDat(j), aryDat(j + 1))
                
                End If
            
            Next j
        
        Next i

End Sub

Public Sub ���Z�b�g()

    With ActiveSheet

        .Pictures.Delete
        .Cells.ClearContents
        .PageSetup.PrintArea = "A1:A3"

    End With

End Sub

Public Sub PDF�����o��()

    Dim ans As String
    Dim fileName As String '�ۑ���t�H���_�p�X&�t�@�C����

    ans = InputBox("�t�@�C�����́H")

    If ans <> "" Then

        fileName = ThisWorkbook.Path&"\"&CStr(ans)&".pdf"

    End If

    With ActiveSheet.PageSetup

        .Zoom = 400
        .Orientation = xlLandscape
        .FitToPageWide = False
        .FitToPageTall = False
        .CenterHorizontally = True
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)

    End With

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
End
