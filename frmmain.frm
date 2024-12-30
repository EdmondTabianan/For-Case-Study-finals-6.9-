VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "Library Borrow/Return a Book System ver 6.9"
   ClientHeight    =   8355
   ClientLeft      =   1560
   ClientTop       =   2010
   ClientWidth     =   17580
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "frmmain.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   17580
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7980
      Width           =   17580
      _ExtentX        =   31009
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "7:17 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "8/11/2024"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   24694
            MinWidth        =   24694
            Text            =   "Library Borrow/Return Book System"
            TextSave        =   "Library Borrow/Return Book System"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   24694
            MinWidth        =   24694
            Text            =   "Developed By: BSIT 1-A Group 5"
            TextSave        =   "Developed By: BSIT 1-A Group 5"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   12960
      Left            =   0
      Picture         =   "frmmain.frx":170AF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   23145
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   8160
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Menu mnumm 
      Caption         =   "Main Menu"
      Begin VB.Menu mnuBLR 
         Caption         =   "Borrow A Book Module"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuh1 
         Caption         =   "-"
      End
      Begin VB.Menu mnureturn 
         Caption         =   "Return Book"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnh2 
         Caption         =   "-"
      End
      Begin VB.Menu mnunew 
         Caption         =   "Add/Remove Book"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuh4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuANA 
         Caption         =   "Add New Admin"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuh5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^{F5}
      End
   End
   Begin VB.Menu mnulogout 
      Caption         =   "Log Out"
      Begin VB.Menu mnuSO 
         Caption         =   "Sign Out"
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuANA_Click()
    sign_up.Show
End Sub

Private Sub mnuBLR_Click()
    frmlborrowingbook.Show
End Sub

Private Sub mnubook_Click()
    frmavailablebooks.Show
    frmavailablebooks.ListView1.Visible = False
    frmavailablebooks.txtselected.Visible = False
End Sub

Private Sub mnuexit_Click()
    If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
    End
End If
End Sub


Private Sub mnunew_Click()
    frmaddbook.Show
End Sub

Private Sub mnureturn_Click()
    frmreturn.Show
End Sub

Private Sub mnuSO_Click()
    frmlogin.Show
    Unload Me
End Sub
