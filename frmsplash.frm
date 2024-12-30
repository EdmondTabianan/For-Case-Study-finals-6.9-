VERSION 5.00
Begin VB.Form frmsplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library Borrow/Return a Book System"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   3120
   End
   Begin VB.Label lblloadingbar 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   10335
   End
   Begin VB.Label lblread 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright 2024"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Developed By: BSIT 1-A Group 5"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version 6.9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Library Borrow/Return a Book System"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   240
      Picture         =   "frmsplash.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y As Integer

Private Sub Form_Load()
    lblloadingbar.Width = X
    X = 10335 / 20
    
End Sub

Private Sub Timer1_Timer()
    lblloadingbar.Width = lblloadingbar.Width + X
    Y = Y + 5
    lblread.Caption = "Now Loading..." & Y & "%"
        If Y >= 100 Then
            MsgBox "Welcome!!! Loading Completed", vbInformation, "Welcome"
            frmlogin.Show
            Unload Me
        End If
End Sub
