VERSION 5.00
Begin VB.Form sign_up 
   BackColor       =   &H00C0E0FF&
   Caption         =   "User Registration Module"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   Picture         =   "sign_up.frx":0000
   ScaleHeight     =   7410
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   11
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   8
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Frame F 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Commands"
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   9615
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Search record"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         TabIndex        =   17
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtsearch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         TabIndex        =   16
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4560
         TabIndex        =   12
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "&Delete "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4560
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6720
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Add new Admin"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Enter Username to search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   4935
      End
   End
   Begin VB.TextBox txtconfirm 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtuser 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1650
      Left            =   7800
      Picture         =   "sign_up.frx":170AF
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Enter "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   14
      Top             =   6120
      Width           =   4575
   End
   Begin VB.Label Label4 
      Caption         =   "Enter Username to search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   13
      Top             =   6120
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000000&
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000000&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "sign_up"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim btn As Integer


Private Sub cmdadd_Click()
    btn = 1
    txtuser.Text = ""
    txtpass.Text = ""
    txtconfirm.Text = ""
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmddelete.Enabled = False
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    cmdclose.Enabled = True
    cmdsearch.Enabled = True
End Sub

Private Sub cmdcancel_Click()
    txtuser.Text = ""
    txtpass.Text = ""
    txtconfirm.Text = ""
    cmdadd.Enabled = True
    cmdedit.Enabled = False
    cmddelete.Enabled = False
    cmdsave.Enabled = False
    cmdcancel.Enabled = True
    cmdclose.Enabled = True
    cmdsearch.Enabled = True
    txtuser.SetFocus
    Set rs = New Recordset
    rs.Open "select *from tbluser", con, adOpenStatic, adLockOptimistic

End Sub

Private Sub cmdclose_Click()
    frmmain.Show
    Unload Me
    rs.Close
End Sub

Private Sub cmddelete_Click()
    If MsgBox("Are you sure you want to delete the current Admin?", vbYesNo + vbQuestion, "Delete") = vbYes Then
        Set rs = New Recordset
            rs.Open "select *from tbluser where adminid=" & AID & "", con, adOpenStatic, adLockOptimistic
            rs.Delete
                MsgBox "Admin Information Deleted!", vbInformation, "Deleted"
            rs.Update
            cmdcancel_Click
    End If
End Sub

Private Sub cmdedit_Click()
    btn = 2
    cmdedit.Enabled = False
    cmddelete.Enabled = False
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    cmdclose.Enabled = True
    cmdsearch.Enabled = True
    cmdsave.Caption = "&Update Record"
End Sub

Private Sub cmdprint_Click()
    Set rs = New Recordset
rs.Open "select *from tbluser where adminid=" & AID & "", con, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.Close
        Set rs = Nothing
        DataReport1.Show
    Else
        Set DataEnvironment1.rsCommand1.DataSource = rs
        DataReport1.Show
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub cmdsave_Click()
Set rs = New Recordset
    If txtuser.Text = "" Or txtpass.Text = "" Or txtconfirm.Text = "" Then
        MsgBox "Field Missing", vbCritical, "Required Input"
        txtuser.SetFocus
    ElseIf txtpass <> txtconfirm Then
        MsgBox "Type Mismatched", vbCritical, "Mismatched"
        txtpass.Text = ""
        txtconfirm.Text = ""
        txtpass.SetFocus
    Else
        If btn = 1 Then
            Set rs = New Recordset
                rs.Open "select *from tbluser", con, adOpenStatic, adLockOptimistic
                rs.AddNew
                    rs!UserName = txtuser
                    rs!Password = txtpass
                    MsgBox "Admin Information Saved", vbInformation, "Saved"
                rs.Update
                cmdcancel_Click
        ElseIf btn = 2 Then
            rs.Open "select *from tbluser where adminid=" & AID & "", con, adOpenStatic, adLockOptimistic
                rs!UserName = txtuser
                rs!Password = txtpass
                rs.Update
                MsgBox "Admin Information Updated", vbInformation, "Updated"
                cmdcancel_Click
                Unload Me
                sign_up.Show
        End If
    End If
End Sub

Private Sub cmdsearch_Click()
    Set rs = New Recordset
    rs.Open "select * from tbluser", con, adOpenStatic, adLockOptimistic
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
                AID = rs!adminid
                If rs!UserName = txtsearch.Text Then
                    MsgBox "Admin Information Found!", vbInformation, "Found"
                    txtuser.Text = rs!UserName
                    txtpass.Text = rs!Password
                    txtconfirm.Text = txtpass.Text
                    cmdedit.Enabled = True
                    cmddelete.Enabled = True
                    Exit Sub
                End If
            rs.MoveNext
            Loop
        Else
            Exit Sub
        End If
        
        MsgBox "Sorry! Admin Information Not Found!", vbInformation, "Not Found"
        txtsearch.Text = ""
        txtsearch.SetFocus
End Sub

Private Sub cmdview_Click()
    frmlist.Show
    Unload Me
End Sub

Private Sub Form_Load()
    txtuser.Text = ""
    txtpass.Text = ""
    txtconfirm.Text = ""
    cmdadd.Enabled = True
    cmdedit.Enabled = False
    cmddelete.Enabled = False
    cmdsave.Enabled = False
    cmdcancel.Enabled = True
    cmdclose.Enabled = True
    cmdsearch.Enabled = True
    Set rs = New Recordset
    rs.Open "select *from tbluser", con, adOpenStatic, adLockOptimistic
End Sub


