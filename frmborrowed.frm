VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmborrowed 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrowed Books"
   ClientHeight    =   6660
   ClientLeft      =   6870
   ClientTop       =   2985
   ClientWidth     =   8115
   FillColor       =   &H00C0FFC0&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H0080FF80&
      Caption         =   "&Clear"
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
      Left            =   5040
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0080FF80&
      Caption         =   "&Back"
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
      Left            =   3120
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H0080FF80&
      Caption         =   "&Select"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtselect 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
   Begin MSComctlLib.ListView book_borrowed 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7223
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
         Text            =   "BorrowID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   10760
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search Book:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmborrowed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset
Dim bookTitle As String
Dim bookID As String

Private Sub book_borrowed_Click()
    txtselect.Text = book_borrowed.selectedItem.ListSubItems(1).Text
End Sub

Private Sub book_borrowed_DblClick()
     AID = book_borrowed.selectedItem.Text
        frmreturn.txttitle.Text = book_borrowed.selectedItem.ListSubItems(1).Text
        frmreturn.lblstatus.Caption = "Book is currently borrowed"
        frmreturn.cmdprintone.Enabled = True
        frmreturn.cmdprintall.Enabled = False
        frmreturn.cmdreturn.Enabled = True
        frmreturn.cmdcancel.Enabled = True
    frmreturn.Show
    Unload Me
End Sub

Private Sub cmdback_Click()
frmreturn.Show
Unload Me
End Sub

Private Sub cmdclear_Click()
    txtselect.Text = ""
    txtselect.SetFocus
End Sub

Private Sub cmdreturn_Click()
    AID = book_borrowed.selectedItem.Text
        frmreturn.txttitle.Text = book_borrowed.selectedItem.ListSubItems(1).Text
        frmreturn.lblstatus.Caption = "Book is currently borrowed"
        frmreturn.cmdprintone.Enabled = True
        frmreturn.cmdprintall.Enabled = False
        frmreturn.cmdreturn.Enabled = True
        frmreturn.cmdcancel.Enabled = True
    frmreturn.Show
    Unload Me
End Sub

Private Sub Form_Load()
    borrowed_book
End Sub

Sub borrowed_book()
    Set rs = New Recordset
        rs.Open "select *from tblborrow", con, adOpenStatic, adLockOptimistic
            If rs.RecordCount = 0 Then
                book_borrowed.ListItems.Clear
            End If
        With book_borrowed
            Do While Not rs.EOF
                .ListItems.Add , , rs!book_id
                .ListItems.Item(.ListItems.Count).ListSubItems.Add , , rs!title

            rs.MoveNext
            Loop
        End With
End Sub

Private Sub txtselect_Change()
    
    Set rs = New Recordset
        rs.Open "select *from tblborrow where title like '" & txtselect.Text & "%" & "'", con, adOpenStatic, adLockOptimistic
            book_borrowed.ListItems.Clear

        With book_borrowed
            Do While Not rs.EOF
                .ListItems.Add , , rs!book_id
                .ListItems.Item(.ListItems.Count).ListSubItems.Add , , rs!title
            rs.MoveNext
            Loop
        End With

End Sub
