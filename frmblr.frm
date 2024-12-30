VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmavailablebooks 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Available Books"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclear 
      Caption         =   "&Clear"
      Height          =   615
      Left            =   5640
      TabIndex        =   9
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton cmdbacks 
      Caption         =   "&BACK"
      Height          =   615
      Left            =   3360
      TabIndex        =   8
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "&BACK"
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   7680
      Width           =   2055
   End
   Begin VB.TextBox txtselected 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   5415
   End
   Begin VB.CommandButton cmdselected 
      Caption         =   "&SELECT"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdselect 
      Caption         =   "&SELECT"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox txtselect 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin MSComctlLib.ListView available_book 
      Height          =   6495
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BOOKID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TITLE"
         Object.Width           =   9701
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6495
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BOOKID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TITLE"
         Object.Width           =   9701
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search Book"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmavailablebooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub available_book_Click()
        txtselect.Text = available_book.selectedItem.ListSubItems(1).Text
End Sub


Private Sub available_book_DblClick()
        AID = available_book.selectedItem.Text
        frmlborrowingbook.txttitle.Text = available_book.selectedItem.ListSubItems(1).Text
        frmlborrowingbook.lblstatus.Caption = "Book is Available"
        frmlborrowingbook.cmdprintall.Enabled = False
        frmlborrowingbook.cmdprintone.Enabled = True
        frmlborrowingbook.cmdborrow.Enabled = True
        frmlborrowingbook.cmdcancel.Enabled = True
    frmlborrowingbook.Show
    Unload Me
End Sub

Private Sub cmdback_Click()
frmlborrowingbook.Show
Unload Me
End Sub

Private Sub cmdbacks_Click()
    frmaddbook.Show
    Unload Me
End Sub

Private Sub cmdclear_Click()
    If txtselect.Visible = True Then
        txtselect.Text = ""
        txtselect.SetFocus
    ElseIf txtselected.Visible = True Then
        txtselected.Text = ""
        txtselected.SetFocus
    End If
End Sub

Private Sub ListView1_Click()
        txtselected.Text = ListView1.selectedItem.ListSubItems(1).Text
End Sub

Private Sub ListView1_DblClick()
    AID = ListView1.selectedItem.Text
        frmaddbook.txttitle.Text = ListView1.selectedItem.ListSubItems(1).Text
        frmaddbook.cmdedit.Enabled = True
        frmaddbook.cmdcancel.Enabled = True
        frmaddbook.cmdremove.Enabled = True
        frmaddbook.Show
        Unload Me
End Sub

Private Sub cmdselect_Click()
    AID = available_book.selectedItem.Text
        frmlborrowingbook.txttitle.Text = available_book.selectedItem.ListSubItems(1).Text
        frmlborrowingbook.lblstatus.Caption = "Book is Available"
        frmlborrowingbook.cmdprintall.Enabled = False
        frmlborrowingbook.cmdprintone.Enabled = True
        frmlborrowingbook.cmdborrow.Enabled = True
        frmlborrowingbook.cmdcancel.Enabled = True
    frmlborrowingbook.Show
    Unload Me
End Sub

'fix code 6.4
Private Sub cmdselected_Click()
    AID = ListView1.selectedItem.Text
        frmaddbook.txttitle.Text = ListView1.selectedItem.ListSubItems(1).Text
        frmaddbook.cmdedit.Enabled = True
        frmaddbook.cmdremove.Enabled = True
        frmaddbook.cmdcancel.Enabled = True
        frmaddbook.cmdsave.Enabled = False
        frmaddbook.cmdadd.Enabled = False
    frmaddbook.Show
    Unload Me
End Sub

Private Sub Form_Load()
    table_book
    add_book
End Sub

Sub table_book()
    Set rs = New Recordset
        rs.Open "select *from tblbooks", con, adOpenStatic, adLockOptimistic
            If rs.RecordCount = 0 Then
                available_book.ListItems.Clear
                
            End If
        With available_book
            Do While Not rs.EOF
                .ListItems.Add , , rs!book_id
                .ListItems.Item(.ListItems.Count).ListSubItems.Add , , rs!title

            rs.MoveNext
            Loop
        End With
End Sub

Sub add_book()
    Set rs = New Recordset
        rs.Open "select *from tblbooks", con, adOpenStatic, adLockOptimistic
            If rs.RecordCount = 0 Then
                ListView1.ListItems.Clear
                 
            End If
        With ListView1
            Do While Not rs.EOF
                .ListItems.Add , , rs!book_id
                .ListItems.Item(.ListItems.Count).ListSubItems.Add , , rs!title

            rs.MoveNext
            Loop
        End With
End Sub


Private Sub txtselect_Change()
    Set rs = New Recordset
        rs.Open "select *from tblbooks where title like '" & txtselect.Text & "%" & "'", con, adOpenStatic, adLockOptimistic
            available_book.ListItems.Clear

        With available_book
            Do While Not rs.EOF
                .ListItems.Add , , rs!book_id
                .ListItems.Item(.ListItems.Count).ListSubItems.Add , , rs!title
            rs.MoveNext
            Loop
        End With
End Sub


Private Sub txtselected_Change()
    Set rs = New Recordset
        rs.Open "select *from tblbooks where title like '" & txtselected.Text & "%" & "'", con, adOpenStatic, adLockOptimistic
            ListView1.ListItems.Clear

        With ListView1
            Do While Not rs.EOF
                .ListItems.Add , , rs!book_id
                .ListItems.Item(.ListItems.Count).ListSubItems.Add , , rs!title
            rs.MoveNext
            Loop
        End With
End Sub
