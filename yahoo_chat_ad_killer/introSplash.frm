VERSION 5.00
Begin VB.Form introSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000D180A&
   BorderStyle     =   0  'None
   Caption         =   "EZ Open"
   ClientHeight    =   3015
   ClientLeft      =   1710
   ClientTop       =   1725
   ClientWidth     =   6540
   ForeColor       =   &H00A2D544&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3015
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   480
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   960
      Top             =   360
   End
   Begin VB.Label lblRaph 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "               by              underground technologies"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   2210
      Visible         =   0   'False
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y a h o o  A d  K i l l e r"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   732
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   5892
   End
   Begin VB.Line L2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      X1              =   480
      X2              =   840
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line L 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   4
      X1              =   5880
      X2              =   6240
      Y1              =   360
      Y2              =   120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   735
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "introSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const WM_SYSCOMMAND = &H112
Dim NoFreeze%

Dim XX1 As Integer
Dim XX2 As Integer
Dim YY1 As Integer
Dim YY2 As Integer
Dim XXX1 As Integer
Dim XXX2 As Integer
Dim YYY1 As Integer
Dim YYY2 As Integer

Dim When As Integer
Dim Start As Boolean
Dim i As Integer

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Const sname = "Y a h o o  A d  K i l l e r" '<-------------- change this to your app title (works better if you leave 1 space between each character
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
        MsgBox "This application is already running.", vbInformation
        End
    End If

Dim v%
Call InitWindow(Me)
YY1 = L.Y1
YY2 = L.Y2
XX1 = L.X1
XX2 = L.X2
YYY1 = L2.Y1
YYY2 = L2.Y2
XXX1 = L2.X1
XXX2 = L2.X2
Start = False
i = 1
lblAppName = ""
lblRaph.Font = 7
'v% = sndPlaySound(App.Path & "\intro.wav", 1): NoFreeze% = DoEvents()

Call InitWindow(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer2.Enabled = False
Timer3.Enabled = False
Set introSplash = Nothing
Unload Me
main.Show
End Sub

Private Sub lblAppName_DblClick()
Unload Me
End Sub

Private Sub Timer2_Timer()
YY2 = YY2 - 100: If YY2 = 600 Then YY2 = 0
YY1 = YY1 + 100: If YY1 = 600 Then YY1 = 0
XX2 = XX2 - 100: If XX2 = 0 Then XX2 = 600
XX1 = XX1 - 100: If XX1 = 0 Then XX1 = 600
L.X1 = XX1
L.X2 = XX2
L.Y1 = YY1
L.Y2 = YY2
End Sub

Private Sub Timer3_Timer()
Dim s As Integer
YYY2 = YYY2 - 100: If YY2 = 0 Then YY2 = 600
YYY1 = YYY1 + 100: If YY1 = 600 Then YY1 = 0
XXX2 = XXX2 + 100: If XX2 = 600 Then XX2 = 0
XXX1 = XXX1 + 100: If XX1 = 600 Then XX1 = 0
L2.X1 = XXX1
L2.X2 = XXX2
L2.Y1 = YYY1
L2.Y2 = YYY2

If L.X1 = 3180 Then
    lblAppName.Visible = True
    Start = True
End If

If Start = True Then
    If L2.X1 = 6480 And L2.Y1 = 6120 Then
        FinishSplash
    ElseIf i = Len(sname) + 1 Then
        Exit Sub
    Else
        A = lblAppName
        b = Mid(sname, i, 1)
        A = A & b
        lblAppName = A
        i = i + 1
    End If
End If
End Sub


Sub FinishSplash()
lblRaph.Visible = True
Shape1.Visible = False
Shape1.Visible = True
lblRaph.Visible = True

lblRaph.FontSize = 10.5
Wait 0.5
lblRaph.FontSize = 10
Wait 0.5
lblRaph.FontSize = 8.5
Wait 0.5
lblRaph.FontSize = 8
Wait 0.5
lblRaph.FontSize = 8.5
Wait 0.5
lblRaph.FontSize = 8
Wait 0.5
lblRaph.FontSize = 3.5
Wait 0.5
lblRaph.FontSize = 3
Wait 0.5
lblRaph.FontSize = 4.5
Wait 0.5
lblRaph.FontSize = 4
Wait 0.5
lblRaph.FontSize = 6.5
Wait 0.5
lblRaph.FontSize = 6
Wait 0.5
lblRaph.FontSize = 7.5
Wait 0.5
lblRaph.FontSize = 7
Wait 0.5
lblRaph.FontSize = 8.5
Wait 0.5
lblRaph.FontSize = 8
Wait 0.5
lblRaph.FontSize = 9.5
Wait 0.5
lblRaph.FontSize = 9
Wait 0.5
lblRaph.FontSize = 8.5
Wait 0.5
lblRaph.FontSize = 8
Wait 0.5
lblRaph.FontSize = 7.5
Wait 0.5
lblRaph.FontSize = 7
Wait 0.5
lblRaph.FontSize = 8.5
Wait 0.5
lblRaph.FontSize = 8
Wait 0.5
lblRaph.FontSize = 7.5
Wait 0.5
lblRaph.FontSize = 7
Wait 0.5
lblRaph.FontSize = 6.5
Wait 0.5
lblRaph.FontSize = 6
Wait 0.5
lblRaph.FontSize = 7.5
Wait 0.5
lblRaph.FontSize = 7
Wait 0.5
lblRaph.FontSize = 9.5
Wait 0.5
lblRaph.FontSize = 9
Wait 0.5
lblRaph.FontSize = 8.5
Wait 0.5
lblRaph.FontSize = 8
Wait 0.5
lblRaph.FontSize = 7.5
Wait 0.5
Shape1.Visible = True
lblRaph.FontSize = 7
lblRaph.FontSize = 7
Wait 0.5
lblRaph.FontSize = 9.5
Wait 0.5
lblRaph.FontSize = 9
Wait 0.5
lblRaph.FontSize = 8.5
Wait 0.5
lblRaph.FontSize = 8
Wait 0.5
lblRaph.FontSize = 7.5
Wait 0.5
Shape1.Visible = True
lblRaph.FontSize = 7
'lblRaph.FontSize = 1.5
Wait 0.5
'lblRaph.FontSize = 1
Wait 0.5
'lblRaph.FontSize = 1.5
Wait 2.5
'lblRaph.FontSize = 7


main.Show
Unload Me
End Sub




Public Function Wait(ByVal TimeToWait As Long)
Dim EndTime As Long
EndTime = GetTickCount + TimeToWait * 1000
Do Until GetTickCount > EndTime
DoEvents
Loop
End Function

