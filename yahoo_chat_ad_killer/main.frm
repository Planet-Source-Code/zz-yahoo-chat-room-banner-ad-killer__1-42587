VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "u-tech's yahoo ad killer - ezbuzzit"
   ClientHeight    =   2355
   ClientLeft      =   150
   ClientTop       =   120
   ClientWidth     =   3510
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
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   345
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3720
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00404040&
      Caption         =   "stay ontop"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      ToolTipText     =   "this will make the program stay on top of all other programs."
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00404040&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      Caption         =   "hover fade"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "this will give your link a faded hover effect."
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "sample"
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   345
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   10
      ToolTipText     =   "click to choose your background color"
      Top             =   1360
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00BEFB31&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C5ED70&
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   9
      ToolTipText     =   "click to choose your text color"
      Top             =   1180
      Width           =   135
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "- ezbuzzit"
      ToolTipText     =   "type here what you want to say after the link."
      Top             =   480
      Width           =   3285
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Text            =   "http://www.geocities.com/underground_technologies"
      ToolTipText     =   "type here the link url"
      Top             =   1200
      Width           =   3045
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "u-tech's"
      ToolTipText     =   "type here the link text"
      Top             =   840
      Width           =   3285
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "ad killed by"
      ToolTipText     =   "type here what you want to say before the link."
      Top             =   120
      Width           =   3285
   End
   Begin VB.CommandButton Command2 
      Caption         =   "preview"
      Height          =   375
      Left            =   1740
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   345
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "main.frx":07F2
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "options"
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "kill ad"
      Height          =   375
      Left            =   55
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label color_Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label color_Label 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Menu mnu_file 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnu_hlp 
         Caption         =   "help"
         Begin VB.Menu mnu_about 
            Caption         =   "about"
         End
      End
      Begin VB.Menu mnu_rthjyr 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_save 
         Caption         =   "save"
      End
      Begin VB.Menu mnu_open 
         Caption         =   "open"
      End
      Begin VB.Menu mnu_teggghds 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_create_reg_key 
         Caption         =   "create key"
      End
      Begin VB.Menu mnu_gfhshgfsgfsghsgshs 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "exit"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LEFT_BUTTON = 1
Const RIGHT_BUTTON = 2
Const MIDDLE_BUTTON = 4


Dim initVals(4) As Long
Private Sub Check1_Click()
If Check1.Value = 1 Then
MsgBox "this sucks because it limits the msg length of the banner! :(", vbInformation
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Call InitWindow(Me)
Else
Call UnInitWindow(Me)
End If
End Sub
Private Sub Command1_Click()
Call txt_Save_Text(Text1, App.Path & "\temp.html")
Load banner_preview
banner_preview.Show , Me
Me.Enabled = False
End Sub
Private Sub Command2_Click()
Dim sysDIR As String
sysDIR = GetSysDir

If Check1.Value = 1 Then
Text6.Text = "<body bgcolor=""" & color_Label2.Caption & """ text=""" & color_Label.Caption & """><SCRIPT language=Javascript src=""file://" & sysDIR & "\u-tech\Fade.js""></SCRIPT><H3><Center>" & Text2.Text & " <a href=" & Text4.Text & " target=_newwin>" & Text3.Text & "</a> " & Text5.Text & "<"
Else
Text6.Text = "<body bgcolor=""" & color_Label2.Caption & """ text=""" & color_Label.Caption & """><H3><Center>" & Text2.Text & " <a href=" & Text4.Text & " target=_newwin>" & Text3.Text & "</a> " & Text5.Text & "<"
End If

Call txt_Save_Text(Text6, App.Path & "\temp.html")
Load banner_preview
banner_preview.Show , Me
Me.Enabled = False
End Sub
Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
      Case LEFT_BUTTON
         PopupMenu main.mnu_file, Command3.Left, Command3.Top + Command3.Height
          Case RIGHT_BUTTON
               End Select
End Sub
Private Sub Command4_Click()
Dim sysDIR As String
    sysDIR = GetSysDir
    
If Check1.Value = 1 Then
Text6.Text = "<body bgcolor=""" & color_Label2.Caption & """ text=""" & color_Label.Caption & """><SCRIPT language=Javascript src=""file://" & sysDIR & "\u-tech\Fade.js""></SCRIPT><H3><Center>" & Text2.Text & " <a href=" & Text4.Text & " target=_newwin>" & Text3.Text & "</a> " & Text5.Text & "<"
SaveSettingString HKEY_CURRENT_USER, "Software\Yahoo\Pager\yurl", "Chat Adurl", "about:<body bgcolor=""" & color_Label2.Caption & """ text=""" & color_Label.Caption & """><SCRIPT language=Javascript src=""file://" & sysDIR & "\u-tech\Fade.js""></SCRIPT><H3><Center>" & Text2.Text & " <a href=" & Text4.Text & " target=_newwin>" & Text3.Text & "</a> " & Text5.Text & "<"
Else
Text6.Text = "<body bgcolor=""" & color_Label2.Caption & """ text=""" & color_Label.Caption & """><SCRIPT language=Javascript src=""file://" & sysDIR & "\u-tech\Fade.js""></SCRIPT><H3><Center>" & Text2.Text & " <a href=" & Text4.Text & " target=_newwin>" & Text3.Text & "</a> " & Text5.Text & "<"
SaveSettingString HKEY_CURRENT_USER, "Software\Yahoo\Pager\yurl", "Chat Adurl", "about:<body bgcolor=""" & color_Label2.Caption & """ text=""" & color_Label.Caption & """><H3><Center>" & Text2.Text & " <a href=" & Text4.Text & " target=_newwin>" & Text3.Text & "</a> " & Text5.Text & "<"
End If

MsgBox "chat room ad has been killed" & vbNewLine & "if you are in a chat room you" & vbNewLine & "will need to leave the room" & vbNewLine & "and come back before" & vbNewLine & "the changes will be seen." & vbNewLine & vbNewLine & "here is what your new banner looks like!", vbInformation: Command2_Click
End Sub
Private Sub Form_Load()
Dim sysDIR As String
sysDIR = GetSysDir
    If FileExists(sysDIR & "\u-tech\Fade.js") = True Then 'Yes it is
        Me.Show
    Else
        MkDir sysDIR & "\u-tech"
        GenFileFromRes 101, "FADE", "FADE", , , sysDIR & "\u-tech\Fade.js"
        GenFileFromRes 102, "STYLE", "STYLE", , , sysDIR & "\u-tech\Style.css"
        Me.Show
    End If
    
If Check3.Value = 1 Then
Call InitWindow(Me)
Else
Call UnInitWindow(Me)
End If

Call Combo_AddFonts(Combo1)
Combo1.Text = "Tahoma"
color_Label.Caption = "#B1F3B6" '#HFFFFF
Picture1.BackColor = &HBEFB31
color_Label2.Caption = "#000000"
Picture2.BackColor = &H0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set main = Nothing
End
End Sub
Private Sub mnu_about_Click()
Me.Enabled = False
u_tech_About.Show , Me
End Sub
Private Sub mnu_create_reg_key_Click()
Dim sysDIR As String
    sysDIR = GetSysDir

Text7.Text = ""
If Check1.Value = 1 Then
Text7.Text = "REGEDIT4" & vbNewLine & vbNewLine & "[HKEY_CURRENT_USER\Software\Yahoo\Pager\yurl]" & vbNewLine & """Finance Disclaimer""=""http://msg.edit.yahoo.com/config/jlb""" & vbNewLine & """Chat Adurl""=""about:<body bgcolor=""" & color_Label2.Caption & """ text=""" & color_Label.Caption & """><SCRIPT language=Javascript src=""file://" & sysDIR & "\u-tech\Fade.js""></SCRIPT><H3><Center>" & Text2.Text & " <a href=" & Text4.Text & " target=_newwin>" & Text3.Text & "</a> " & Text5.Text & "<"
Call txt_Save_Text(Text7, App.Path & "\temp.reg")
Else
Text7.Text = "REGEDIT4" & vbNewLine & vbNewLine & "[HKEY_CURRENT_USER\Software\Yahoo\Pager\yurl]" & vbNewLine & """Finance Disclaimer""=""http://msg.edit.yahoo.com/config/jlb""" & vbNewLine & """Chat Adurl""=""about:<body bgcolor=""" & color_Label2.Caption & """ text=""" & color_Label.Caption & """><H3><Center>" & Text2.Text & " <a href=" & Text4.Text & " target=_newwin>" & Text3.Text & "</a> " & Text5.Text & "<"
Call txt_Save_Text(Text7, App.Path & "\temp.reg")
End If
End Sub
Private Sub mnu_exit_Click()
Unload Me
End Sub
Private Sub mnu_open_Click()
Dim l036A As Variant
Dim l036E As Variant
Dim gv0006, l0078$
Dim MYSTR As String, FilePath2 As String, textz As String, A As String
    CommonDialog1.Filename = ""
    CommonDialog1.CancelError = True
    CommonDialog1.DefaultExt = "ini"
    CommonDialog1.DialogTitle = "Select ID Files..."
    CommonDialog1.Filter = "INI Files (*.ini)|*.ini|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo L3CC96
    CommonDialog1.Action = 1
    gv0006 = CommonDialog1.Filename
    
    If gv0006 = "" Then Exit Sub

Text2 = Get_From_INI("banner", "pre text", CommonDialog1.Filename)
Text5 = Get_From_INI("banner", "post text", CommonDialog1.Filename)
Text3 = Get_From_INI("banner", "link text", CommonDialog1.Filename)
Text4 = Get_From_INI("banner", "url text", CommonDialog1.Filename)

Exit Sub
L3CC96:

End Sub
Private Sub mnu_save_Click()
Dim R%
Dim SaveList As Long

Dim l036A As Variant
Dim l036E As Variant
Dim gv0006, l0078$
Dim MYSTR As String, FilePath2 As String, textz As String, A As String

On Error Resume Next
CommonDialog1.Filename = ""
    CommonDialog1.CancelError = True
    CommonDialog1.DefaultExt = "ini"
    CommonDialog1.DialogTitle = "Select ID Files..."
    CommonDialog1.Filter = "INI Files (*.ini)|*.ini|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo L3CC96
    CommonDialog1.Action = 1
    gv0006 = CommonDialog1.Filename
If gv0006 = "" Then Exit Sub

R% = WritePrivateProfileString("banner", "pre text", Text2.Text, gv0006)
R% = WritePrivateProfileString("banner", "post text", Text5.Text, gv0006)
R% = WritePrivateProfileString("banner", "link text", Text3.Text, gv0006)
R% = WritePrivateProfileString("banner", "url text", Text4.Text, gv0006)
Exit Sub
L3CC96:
End Sub
Private Sub Picture1_Click()
  Load color_picker
  color_picker.picView.BackColor = Picture1.BackColor
  color_picker.Show , Me
  Me.Enabled = False
  color_picker.Label1.Caption = "false"
End Sub
Private Sub Picture2_Click()
  Load color_picker
  color_picker.picView.BackColor = Picture2.BackColor
  color_picker.Show , Me
  Me.Enabled = False
  color_picker.Label1.Caption = "true"
End Sub
