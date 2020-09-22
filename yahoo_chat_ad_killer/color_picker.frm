VERSION 5.00
Begin VB.Form color_picker 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picView 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   120
      ScaleHeight     =   1440
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   480
      MouseIcon       =   "color_picker.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "color_picker.frx":0152
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "false"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lbCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RGB:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lbCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HEX:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Index           =   2
      Left            =   1320
      TabIndex        =   6
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Label lbCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HSL:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label lbRgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   120
      Width           =   45
   End
   Begin VB.Label lbHex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   480
      Width           =   45
   End
   Begin VB.Label lbVB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   840
      Width           =   45
   End
   Begin VB.Label lbHsl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1200
      Width           =   45
   End
   Begin VB.Line zLine 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   11
      X1              =   1320
      X2              =   3240
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line zLine 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   8
      X1              =   1320
      X2              =   3240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line zLine 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   9
      X1              =   1320
      X2              =   3240
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line zLine 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   4
      X1              =   1320
      X2              =   3240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line zLine 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   5
      X1              =   1320
      X2              =   3240
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line zLine 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   6
      X1              =   1320
      X2              =   3240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line zLine 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   7
      X1              =   1320
      X2              =   3240
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line zLine 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   10
      X1              =   1320
      X2              =   3240
      Y1              =   1455
      Y2              =   1455
   End
End
Attribute VB_Name = "color_picker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim my_color As String
Dim my_color2 As String




Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True
Set color_picker = Nothing
End Sub


Private Sub picColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    main.picSelect.BackColor = picColor.Point(X, Y)
'   ' showSelectedCol selIndex, picSelect.BackColor
'    lSelected = main.picSelect.BackColor
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    picView.BackColor = picColor.Point(X, Y)
    UpdateColor picView.BackColor
End Sub


Public Sub UpdateColor(lColor As Long)
    color_picker.lbRgb.Caption = rgbRed(lColor) & "," & rgbGreen(lColor) & "," & rgbBlue(lColor)
    color_picker.lbHex.Caption = "#" & HexRGB(lColor)
    color_picker.lbVB.Caption = "&&H" & Hex$(lColor)
    color_picker.lbHsl.Caption = RGBtoHSL(lColor).Hue & "," & RGBtoHSL(lColor).Sat & "," & RGBtoHSL(lColor).Lum
    lSelected = lColor
End Sub


Private Sub picColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = picColor.Point(X, Y)

If Label1.Caption = "false" Then Call set_top_color
If Label1.Caption = "true" Then Call set_bottom_color
End Sub



Public Sub set_top_color()
'On Error Resume Next
    main.Picture1.BackColor = Label2.Caption
    UpdateColor main.Picture1.BackColor
    my_color = color_picker.lbHex.Caption
    my_color2 = picView.BackColor
    main.color_Label = my_color
    'my_color = "#" & main.Picture1.BackColor
'r% = WritePrivateProfileString("ascii", "solid color", my_color, App.Path & "\FaQ.ini")
'r% = WritePrivateProfileString("ascii", "bg color", my_color2, App.Path & "\FaQ.ini")

Unload Me

End Sub
Public Sub set_bottom_color()
'On Error Resume Next
    main.Picture2.BackColor = Label2.Caption
    UpdateColor main.Picture2.BackColor
    my_color = color_picker.lbHex.Caption
    my_color2 = picView.BackColor
    main.color_Label2 = my_color
    'my_color = "#" & main.Picture2.BackColor
'r% = WritePrivateProfileString("ascii", "solid color", my_color, App.Path & "\FaQ.ini")
'r% = WritePrivateProfileString("ascii", "bg color", my_color2, App.Path & "\FaQ.ini")
Unload Me
End Sub

