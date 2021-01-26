Attribute VB_Name = "Module1"
Option Explicit

Sub subName()
    '// add declarations
    On Error GoTo catchError
exitSub:
    Exit Sub
catchError:
    '// add error handling
    GoTo exitSub
End Sub