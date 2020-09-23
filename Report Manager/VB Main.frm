VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6 Reports & Queries"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Open_Queries 
      Caption         =   "Open Access97 Query Manager"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   3555
   End
   Begin VB.CommandButton Open_Reports 
      Caption         =   "Open Access97 Report Manager"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   3555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Open_Queries_Click()
    Shell "C:\Program Files\Microsoft Office\Office\Msaccess.exe """ & App.Path & "\Report~1.mdb"" /x Queries", vbMinimizedFocus
End Sub

Private Sub Open_Reports_Click()
    Shell "C:\Program Files\Microsoft Office\Office\Msaccess.exe """ & App.Path & "\Report~1.mdb"" /x Reports", vbMinimizedFocus
End Sub
