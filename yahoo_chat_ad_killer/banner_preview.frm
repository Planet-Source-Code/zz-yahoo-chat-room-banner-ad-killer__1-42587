VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form banner_preview 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "banner ad preview"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      ExtentX         =   14843
      ExtentY         =   1720
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "banner_preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call InitWindow(Me)
WebBrowser1.Navigate App.Path & "\temp.html"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call WipeFileClean(App.Path & "\temp.html")
main.Enabled = True
Set banner_preview = Nothing
End Sub

