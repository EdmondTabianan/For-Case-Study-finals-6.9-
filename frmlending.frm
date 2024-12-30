VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlborrowingbook 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrow a Book Module"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10110
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
      Left            =   3720
      TabIndex        =   12
      Top             =   4320
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
      Left            =   1200
      TabIndex        =   11
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton cmdhome 
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
      Left            =   7320
      TabIndex        =   10
      Top             =   4320
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   2535
      Left            =   0
      TabIndex        =   7
      Top             =   5520
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   4471
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:10 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "8/8/2024"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Borrow A  Book"
            TextSave        =   "Borrow A  Book"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   70556
            MinWidth        =   70556
            Picture         =   "frmlending.frx":0000
         EndProperty
      EndProperty
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
      Left            =   3840
      TabIndex        =   6
      Top             =   480
      Width           =   3135
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
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Width           =   6375
   End
   Begin VB.CommandButton cmdborrow 
      Caption         =   "&Borrow"
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
      Left            =   7320
      TabIndex        =   3
      Top             =   3360
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
      Left            =   7320
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "&VIEW AVAILABLE"
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
      Left            =   7320
      TabIndex        =   1
      Top             =   1440
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
      Left            =   7320
      TabIndex        =   0
      Top             =   2400
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
      Left            =   720
      TabIndex        =   9
      Top             =   3120
      Width           =   6255
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
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
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
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "frmlborrowingbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset

Private Sub cmdborrow_Click()
    Set rs = New Recordset
    ' Check if the title input is empty
If txttitle.Text = "" Then
    MsgBox "Required Input", vbInformation, "Required"
    txttitle.SetFocus
Else
    ' Open the Recordset and check if the book title exists in tblbooks
    rs.Open "SELECT * FROM tblbooks WHERE title = '" & txttitle.Text & "'", con, adOpenStatic, adLockOptimistic
    
    ' Check if the recordset is empty (no book found)
    If rs.EOF Then
        MsgBox "Book not found", vbInformation, "Not Found"
    Else
        ' Insert the book title into tblborrow table
        con.Execute "INSERT INTO tblborrow (title) VALUES ('" & txttitle.Text & "')"
        MsgBox "Book Borrowed Successfully", vbInformation, "Success"
        rs.Delete
        cmdcancel_Click
        Unload Me
        frmlborrowingbook.Show
    End If
    rs.Close
End If
    
            
    
End Sub

Private Sub cmdcancel_Click()
    txttitle.Text = ""
    lblstatus.Caption = ""
    cmdsearch.Enabled = True
    cmdprintone.Enabled = False
    cmdprintall.Enabled = True
    cmdborrow.Enabled = False
    cmdcancel.Enabled = False
End Sub

Private Sub cmdhome_Click()
    Unload Me
End Sub

Private Sub cmdprintone_Click()
    Set rs = New Recordset
rs.Open "select *from tblbooks where book_id=" & AID & "", con, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.Close
        Set rs = Nothing
        available.Show
    Else
        Set book_available.rsCommand1.DataSource = rs
        available.Show
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub cmdprintall_Click()
    Set rs = New Recordset
rs.Open "select *from tblbooks", con, adOpenStatic, adLockOptimistic
    Set book_available.rsCommand1.DataSource = rs
    available.Show
End Sub

Private Sub cmdsearch_Click()
    Set rs = New Recordset
    rs.Open "select * from tblbooks", con, adOpenStatic, adLockOptimistic
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
                AID = rs!book_id
                If rs!title = txtsearch.Text Then
                    MsgBox "Book Available", vbInformation, "Available"
                    txttitle.Text = rs!title
                    txtsearch.Text = ""
                    lblstatus.Caption = "Book is Available"
                    cmdprintone.Enabled = True
                    cmdprintall.Enabled = False
                    cmdborrow.Enabled = True
                    cmdcancel.Enabled = True
                    Exit Sub
                End If
            rs.MoveNext
            Loop
        Else
            Exit Sub
        End If
        
        lblstatus.Caption = "Book not found or not yet returned"
        txtsearch.Text = ""
        txtsearch.SetFocus
End Sub

Private Sub cmdview_Click()
    frmavailablebooks.Show
    Unload Me
    frmavailablebooks.ListView1.Visible = False
    frmavailablebooks.txtselected.Visible = False
    frmavailablebooks.cmdbacks.Visible = False
    frmavailablebooks.cmdback.Visible = True
End Sub


Private Sub Form_Load()
    cmdcancel.Enabled = False
    cmdprintone.Enabled = False
    cmdprintall.Enabled = True
    cmdborrow.Enabled = False
    Set rs = New Recordset
    rs.Open "select *from tblbooks", con, adOpenStatic, adLockOptimistic
    cmdview.Caption = "View Available (" & rs.RecordCount & ")"
End Sub

