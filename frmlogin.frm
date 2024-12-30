VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Login"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8730
   FillColor       =   &H0080FF80&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   8730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "show password"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFFF80&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4800
      MaskColor       =   &H00FFC0FF&
      Picture         =   "frmlogin.frx":0000
      TabIndex        =   6
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "&Login"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtuser 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   5760
      Picture         =   "frmlogin.frx":170AF
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim try As Integer

Private Sub Check1_Click()
    If txtpass.PasswordChar = "*" Then
        txtpass.PasswordChar = ""
    Else
        txtpass.PasswordChar = "*"
    End If
End Sub

Private Sub cmdcancel_Click()
    txtuser.Text = ""
    txtpass.Text = ""
    txtuser.SetFocus
End Sub

Private Sub cmdexit_Click()
    If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
    End
End If
End Sub

Private Sub cmdlogin_Click()
    Set rs = New Recordset
        If txtuser.Text = "" Then
            MsgBox "Field Missing", vbCritical, "Invalid Input"
            txtuser.SetFocus
        ElseIf txtpass.Text = "" Then
            MsgBox "Field Missing", vbCritical, "Invalid Input"
            txtpass.SetFocus
        Else
            Set rs = New Recordset
  
  
  
  rs.Open "select *from tbluser where username='" & txtuser & "' and password ='" & txtpass & "'", con, adOpenStatic, adLockOptimistic
    If rs.RecordCount <> 0 Then
        If rs!UserName = txtuser And rs!Password = txtpass Then
            MsgBox "Welcome", vbInformation, "Welcome"
            frmmain.Show
            Unload Me
        Else
            MsgBox "Access Denied", vbInformation, "Invalid"
            txtuser = ""
            txtpass = ""
            txtuser.SetFocus
                try = try + 1
                If try = 5 Then
                MsgBox "Maximum attempts has been reached", vbInformation, "Program Terminated"
                End
            End If
        End If
    Else
            MsgBox "Access Denied", vbInformation, "Invalid"
            txtuser = ""
            txtpass = ""
            txtuser.SetFocus
                try = try + 1
                If try = 5 Then
                MsgBox "Maximum attempts has been reached", vbInformation, "Program Terminated"
                End
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    If cmdlogin.Default = False Then
        cmdlogin.Default = True
    End If
End Sub

