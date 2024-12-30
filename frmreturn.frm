VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreturn 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Module"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   12120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprintall 
      Caption         =   "&PRINT ALL"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4440
      TabIndex        =   12
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmdprintone 
      Caption         =   "&PRINT SELECTED"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1680
      TabIndex        =   11
      Top             =   4200
      Width           =   2415
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   2175
      Left            =   0
      TabIndex        =   10
      Top             =   4920
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   3836
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
            Text            =   "Return module"
            TextSave        =   "Return module"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
            Picture         =   "frmreturn.frx":0000
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Home"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   9
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "&VIEW BORROWED"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "&Return Book"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txttitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   2400
      Width           =   6495
   End
   Begin VB.TextBox txtsearch 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label rrer 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search Book"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Book Title:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblstatus 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   3240
      Width           =   6495
   End
End
Attribute VB_Name = "frmreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset
Dim title As String

Private Sub cmdcancel_Click()
    txttitle.Text = ""
    lblstatus.Caption = ""
    cmdprintone.Enabled = False
    cmdprintall.Enabled = True
    cmdreturn.Enabled = False
    cmdcancel.Enabled = False
End Sub

Private Sub cmdprint_Click()
    Set rs = New Recordset
rs.Open "select *from tblborrow where book_id=" & AID & "", con, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.Close
        Set rs = Nothing
        borrowed.Show
    Else
        Set book_borrow.rsCommand2.DataSource = rs
        borrowed.Show
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub cmdprintselected_Click()
    Set rs = New Recordset
rs.Open "select *from tblborrow", con, adOpenStatic, adLockOptimistic
    Set book_borrow.rsCommand2.DataSource = rs
    borrowed.Show
End Sub

Private Sub cmdprintall_Click()
    Set rs = New Recordset
rs.Open "select *from tblborrow", con, adOpenStatic, adLockOptimistic
    Set book_borrow.rsCommand2.DataSource = rs
    borrowed.Show
End Sub

Private Sub cmdprintone_Click()
    Set rs = New Recordset
rs.Open "select *from tblborrow where book_id=" & AID & "", con, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.Close
        Set rs = Nothing
        DataReportr1.Show
    Else
        Set DataEnvironment1.rsCommand1.DataSource = rs
        DataReport1.Show
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub cmdreturn_Click()
    Set rs = New Recordset
    ' Check if the title input is empty
If txttitle.Text = "" Then
    MsgBox "Required Input", vbInformation, "Required"
    txttitle.SetFocus
Else
    ' Open the Recordset and check if the book title exists in tblbooks
    rs.Open "SELECT * FROM tblborrow WHERE title = '" & txttitle.Text & "'", con, adOpenStatic, adLockOptimistic
    
    ' Check if the recordset is empty (no book found)
    If rs.EOF Then
        MsgBox "Book not found", vbInformation, "Not Found"
    Else
        ' Insert the book title into tblborrow table
        con.Execute "INSERT INTO tblbooks(title) VALUES ('" & txttitle.Text & "')"
        MsgBox "Book Returned successfully", vbInformation, "Returned"
        'lblstatus.Caption =
        rs.Delete
        cmdcancel_Click
        Unload Me
        frmreturn.Show
    End If
End If
    'rs.Close
End Sub

Private Sub cmdsearch_Click()
    Set rs = New Recordset
    rs.Open "select * from tblborrow", con, adOpenStatic, adLockOptimistic
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
                AID = rs!book_id
                If rs!title = txtsearch.Text Then
                    MsgBox "Book found", vbInformation, "found"
                    txttitle.Text = rs!title
                    txtsearch.Text = ""
                    lblstatus.Caption = "Book is currently borrowed"
                    cmdprintone.Enabled = True
                    cmdprintall.Enabled = False
                    cmdreturn.Enabled = True
                    cmdcancel.Enabled = True
                    Exit Sub
                End If
            rs.MoveNext
            Loop
        Else
            Exit Sub
        End If
        
        lblstatus.Caption = "Book not found or already returned"
        txtsearch.Text = ""
        txtsearch.SetFocus
End Sub

Private Sub cmdview_Click()
    frmborrowed.Show
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmdprintone.Enabled = False
    cmdprintall.Enabled = True
    cmdreturn.Enabled = False
    cmdcancel.Enabled = False
    Set rs = New Recordset
    rs.Open "select *from tblborrow", con, adOpenStatic, adLockOptimistic
    cmdview.Caption = "View Borrowed (" & rs.RecordCount & ")"
End Sub

