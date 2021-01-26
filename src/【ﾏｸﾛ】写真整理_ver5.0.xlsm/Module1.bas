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
    fName = Application.GetOpenFilename(Filefilter:="画像ファイル, *.jpg; *.png; *.gif; *.tif;*.bmp", MultiSelect:=True)

    If IsArray(fName) <> True Then

        MsgBox "キャンセルしました"
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
                        '下記いずれかをコメントアウト
                        .ShapeRange.Height = ActiveCell.Height 'セルの高さにリサイズ
                        '.ShapeRange.Width = ActiveCell.Width 'セルの幅にリサイズ

                        With ActiveCell.Offset(0, 1)

                            .Value = fName(i) 'ファイル名の書き込み

                            If i = 1 Then

                                .Font.ColorIndex = 3

                            End If
                        
                        End With
                    
                    End With
                
                ActiveCell.Offset(1, 0).Activate
                Application.StatusBar = "処理中：" & i & "/" & UBound(fName) & "枚目"
                
            Next i

    End If

    With Application

        .StatusBar = False
        .ScreenUpdating = True

    End With

    Set Pict = Nothing

    ActiveSheet.PageSetup.PrintArea = Range(Cells(1, 1), Cells(object + UBound(fName) - LBound(fName), 1)).Address
    MsgBox i - 1 & "枚の画像を挿入しました", vbInformation

End Sub

'値の入れ替え
Public Sub Swap(ByRef Dat1 As Variant, ByRef Dat2 As Variant)

    Dim varBuf As Variant
    varBuf = Dat1
    Dat1 = Dat2
    Dat2 = varBuf

End Sub

'配列のバブルソート
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

Public Sub リセット()

    With ActiveSheet

        .Pictures.Delete
        .Cells.ClearContents
        .PageSetup.PrintArea = "A1:A3"

    End With

End Sub

Public Sub PDF書き出し()

    Dim ans As String
    Dim fileName As String '保存先フォルダパス&ファイル名

    ans = InputBox("ファイル名は？")

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
