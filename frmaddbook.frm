VERSION 5.00
Begin VB.Form frmaddbook 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Add/Remove Book"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsearchbook 
      Caption         =   "Search &Book"
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
      Left            =   7200
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtsearch 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   7
      Top             =   4680
      Width           =   3975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save book"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "&Home"
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
      Left            =   7080
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add Book"
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
      Left            =   3000
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txttitle 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   0
      Top             =   840
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   8895
      Begin VB.CommandButton cmdview 
         Caption         =   "&view Availble book"
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
         Left            =   5640
         TabIndex        =   12
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit Book Title"
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
         Left            =   3720
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdremove 
         Caption         =   "&Remove Book"
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
         Left            =   1680
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search Book"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter new book"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmaddbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rs As New Recordset
Dim bt As Integer

Private Sub cmdadd_Click()
    bt = 1
    txttitle.Text = ""
    txttitle.SetFocus
    cmdadd.Enabled = False
    cmdcancel.Enabled = True
    cmdsave.Enabled = True
    cmdhome.Enabled = True
    cmdsave.Caption = "&Save Book"
End Sub

Private Sub cmdcancel_Click()
    txttitle.Text = ""
    cmdsave.Enabled = False
    cmdedit.Enabled = False
    cmdremove.Enabled = False
    cmdadd.Enabled = True
    cmdview.Enabled = True
    txttitle.SetFocus
    cmdsave.Caption = "&Save Book"
End Sub

Private Sub cmdedit_Click()
    bt = 2
    cmdadd.Enabled = True
    cmdedit.Enabled = False
    cmdhome.Enabled = True
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    cmdadd.Enabled = False
    cmdview.Enabled = False
    cmdsave.Caption = "&Update Book Title"
End Sub

Private Sub cmdhome_Click()
    frmmain.Show
    Unload Me
End Sub

Private Sub cmdremove_Click()
    If MsgBox("Are you sure you want to remove this Book?", vbYesNo + vbQuestion, "Delete") = vbYes Then
        Set rs = New Recordset
        rs.Open "select *from tblbooks where book_id=" & AID & "", con, adOpenStatic, adLockOptimistic
        rs.Delete
            MsgBox "Book Successfully Removed!", vbInformation, "Removed"
        rs.Update
        cmdcancel_Click
        Unload Me
        frmaddbook.Show
    End If
End Sub

Private Sub cmdsave_Click()
    Set rs = New Recordset
    If txttitle.Text = "" Then
        MsgBox "Field Missing", vbCritical, "Required Input"
        txttitle.SetFocus
    Else
        If bt = 1 Then
            Set rs = New Recordset
                rs.Open "select *from tblbooks", con, adOpenStatic, adLockOptimistic
                rs.AddNew
                    rs!title = txttitle
                    MsgBox "Book Saved", vbInformation, "saved"
                rs.Update
                cmdcancel_Click
                Unload Me
                frmaddbook.Show
        ElseIf bt = 2 Then
        rs.Open "select *from tblbooks where book_id = " & AID & "", con, adOpenStatic, adLockOptimistic
                    rs!title = txttitle
                rs.Update
                    MsgBox "Title Updated", vbInformation, "Updated"
                    cmdadd.Enabled = True
                    cmdview.Enabled = True
                    cmdcancel_Click
        End If
    End If
End Sub

Private Sub cmdview_Click()
    frmavailablebooks.Show
    frmavailablebooks.cmdselect.Visible = False
    frmavailablebooks.cmdselected.Visible = True
    Unload Me
    frmavailablebooks.available_book.Visible = False
    frmavailablebooks.ListView1.Visible = True
    frmavailablebooks.txtselect.Visible = False
    frmavailablebooks.cmdback.Visible = False
    frmavailablebooks.cmdbacks.Visible = True
End Sub

Private Sub cmdsearchbook_Click()
     Set rs = New Recordset
    rs.Open "select *from tblbooks", con, adOpenStatic, adLockOptimistic
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
            AID = rs!book_id
            If rs!title = txtsearch.Text Then
                MsgBox "Book has been Found!", vbInformation, "Found"
                txttitle.Text = rs!title
                cmdedit.Enabled = True
                cmdremove.Enabled = True
                cmdcancel.Enabled = True
                txtsearch.Text = ""
                Exit Sub
            End If
        rs.MoveNext
        Loop
    Else
        Exit Sub
    End If

    MsgBox "Sorry! Book Not Found!", vbInformation, "Not Found"
    txtsearch.Text = ""
    txtsearch.SetFocus
    frmavailablebooks.cmdselect.Visible = False
End Sub

Private Sub Form_Load()
    cmdadd.Enabled = True
    cmdview.Enabled = True
    cmdcancel.Enabled = True
    cmdsave.Enabled = False
    cmdhome.Enabled = True
    cmdremove.Enabled = False
    cmdedit.Enabled = False
    Set rs = New Recordset
    rs.Open "select *from tblbooks", con, adOpenStatic, adLockOptimistic
    cmdview.Caption = "View Records (" & rs.RecordCount & ")"
End Sub
