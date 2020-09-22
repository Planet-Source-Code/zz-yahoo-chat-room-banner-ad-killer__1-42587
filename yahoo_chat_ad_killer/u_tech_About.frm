VERSION 5.00
Begin VB.Form u_tech_About 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "u_tech_About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraURL 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Left            =   2700
      TabIndex        =   7
      Top             =   1680
      Width           =   2655
      Begin VB.Label lblURL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "click to visit our website!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "u_tech_About.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   0
         Width           =   2325
      End
   End
   Begin VB.PictureBox picBackgroundBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   120
      Picture         =   "u_tech_About.frx":015E
      ScaleHeight     =   3525
      ScaleWidth      =   2310
      TabIndex        =   0
      Top             =   120
      Width           =   2370
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   330
      Left            =   4320
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblWarnTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warning:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2580
      TabIndex        =   9
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lblRelease 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      ForeColor       =   &H0080C0FF&
      Height          =   240
      Left            =   60
      TabIndex        =   6
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label lblMail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by U-Tech Inc."
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1380
      Width           =   2775
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warning"
      ForeColor       =   &H0080C0FF&
      Height          =   1575
      Left            =   2580
      TabIndex        =   4
      Top             =   2280
      Width           =   2835
   End
   Begin VB.Label lblRights 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      ForeColor       =   &H0080C0FF&
      Height          =   735
      Left            =   2640
      TabIndex        =   3
      Top             =   540
      Width           =   2775
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U-Tech About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "u_tech_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bInsideURL As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Public Sub Start()
    lblName = App.ProductName
    lblRelease = "Release v" & App.Major & "." & _
        App.Minor & " (Build " & App.Revision & ")"
    Show vbModal
End Sub

Private Sub Form_Load()
Call InitWindow(Me)
    Caption = "About " & App.ProductName
    lblRights = "Copyright © 2001, 2003" & vbCrLf & "U-Tech Inc." & vbCrLf & "All rights reserved."
    lblWarning = "This computer program is protected by copyright law and international treaties. Unauthorized reproduction or distribution of this program, or any portion of it, may result in severe civil and criminal penalties, and will be prosecuted to the maximum extent possible under the law."
    lblRelease.Caption = "Release Version " & App.Major & "." & App.Minor & "." & App.Revision
    SetWindowPos Me.hWnd, (-1), 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True
Set u_tech_About = Nothing
End Sub

Private Sub fraURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblURL_Click
End Sub

Private Sub fraURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_bInsideURL Then
        If ((X < 0) Or (Y < 0) Or (Y > fraURL.Height) Or _
            (X > fraURL.Width)) Then
            ReleaseCapture
            lblURL.Font.Underline = False
            lblURL.ForeColor = &HC00000
            m_bInsideURL = False
        End If
    End If
End Sub

Private Sub lblURL_Click()
    Screen.MousePointer = vbHourglass
    DoEvents
    fraURL_MouseMove 0, 0, -1, 0
    Call ShellExecute(Me.hWnd, vbNullString, "http://www.geocities.com/underground_technologies", vbNullString, "C:\", 0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_bInsideURL) Then
        ReleaseCapture
        SetCapture fraURL.hWnd
        lblURL.Font.Underline = True
        lblURL.ForeColor = vbBlue
        m_bInsideURL = True
        Set fraURL.MouseIcon = lblURL.MouseIcon
        fraURL.MousePointer = vbCustom
    End If
End Sub

Private Sub cmdOK_Click()
   Unload Me
End Sub

