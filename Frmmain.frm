VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introduction to HTML Help"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   HelpContextID   =   1001
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.FileListBox File1 
      Height          =   3405
      HelpContextID   =   1003
      Left            =   4545
      TabIndex        =   1
      Top             =   120
      WhatsThisHelpID =   30001
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   105
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   30000
      Width           =   4395
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetHelpFile()
    App.HelpFile = AppDir & "myHelp.chm::popups.txt"
End Sub

Private Function AppDir() As String
    '//If AppDir = NOT the root of the HD  add "\", else don't add the slash
    '//(This sub makes sure an "\" is added to the app.path)
    If Right$(App.Path, 1) <> "\" Then
        AppDir = App.Path & "\"
    Else
        AppDir = App.Path
    End If
End Function

Private Sub Form_Load()
    SetHelpFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub
