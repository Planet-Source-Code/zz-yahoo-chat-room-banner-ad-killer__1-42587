Attribute VB_Name = "Module1"
Option Explicit
Dim X
Sub Combo_AddFonts(Cmb As ComboBox)
    Dim i
    For i = 0 To Screen.FontCount - 1
        Cmb.AddItem Screen.Fonts(i)
    Next i
End Sub

Sub txt_Save_Text(Txt As Control, FilePath3 As String)
    Open FilePath3$ For Output As #1
        Print #1, Txt
    Close 1
End Sub
Sub txt_Load_Text(Txt As TextBox, FilePath As String)
On Error GoTo pothead
    Dim MYSTR As String, FilePath2 As String, textz As String, A As String
    Open FilePath For Input As #1
    Do While Not EOF(1)
    Line Input #1, A$
        textz$ = textz$ + A$ + Chr$(13) + Chr$(10)
        Loop
        Txt = textz$
    Close #1
    Exit Sub
pothead:
main.Text5 = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "       this user has chosen not to have a profile."
    Exit Sub
End Sub
