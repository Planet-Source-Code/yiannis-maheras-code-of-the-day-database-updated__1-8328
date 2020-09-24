VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{57A90BB1-9F57-11D3-B479-A0A072A969C6}#7.0#0"; "VBWHYPERLINK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Planet Source Code Database"
   ClientHeight    =   8436
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12192
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8436
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   684
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   1207
      ButtonWidth     =   1588
      ButtonHeight    =   1164
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Insert"
            Key             =   "cInsert"
            Object.ToolTipText     =   "Saves current record"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete"
            Key             =   "cDelete"
            Object.ToolTipText     =   "Delete current record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import"
            Key             =   "cImport"
            Object.ToolTipText     =   "Importing new code"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Clear"
            Key             =   "cClear"
            Object.ToolTipText     =   "Clears the contents of the fields"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "cSearch"
            Object.ToolTipText     =   "Search for specific code"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Bookmark"
            Key             =   "cBookMark"
            Object.ToolTipText     =   "Bookmarks the current record"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E-mail"
            Key             =   "cMail"
            Object.ToolTipText     =   "Checks for PSC mails"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Recent"
            Key             =   "cRecent"
            Object.ToolTipText     =   "Show the last update"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "cPref"
            Object.ToolTipText     =   "Program options"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Old e-mails"
            Key             =   "cOldMail"
            Object.ToolTipText     =   "Display older e-mail's"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "cAbout"
            Object.ToolTipText     =   "About the program"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "cExit"
            Object.ToolTipText     =   "Ends the application"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   3720
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":235E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3516
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":640A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bookmarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2052
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   12012
      Begin MSComctlLib.ListView cBookmarksList 
         Height          =   1692
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   11772
         _ExtentX        =   20765
         _ExtentY        =   2985
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Level"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Category"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5412
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   4332
      _ExtentX        =   7641
      _ExtentY        =   9546
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame cFrame 
      Height          =   5532
      Left            =   4560
      TabIndex        =   0
      Top             =   720
      Width           =   7572
      Begin VB.TextBox cMailDate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   22
         Top             =   5040
         Width           =   1812
      End
      Begin VB.CheckBox cReviewed 
         Caption         =   "Shows if you have visit already this page"
         Height          =   972
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   1332
      End
      Begin Hyperlink.vbwHyperlink cURL 
         Height          =   216
         Left            =   1440
         TabIndex        =   16
         Top             =   3600
         Width           =   1512
         _ExtentX        =   2667
         _ExtentY        =   381
         HoverColour     =   16711680
         Caption         =   "Planet Source Code"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin VB.TextBox cTitle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   6012
      End
      Begin VB.TextBox cCategory 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   3612
      End
      Begin VB.TextBox cLevel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   4
         Top             =   1200
         Width           =   3612
      End
      Begin VB.TextBox cDes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1812
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1680
         Width           =   6012
      End
      Begin VB.TextBox cCompatibility 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   2
         Top             =   4080
         Width           =   5412
      End
      Begin VB.TextBox cSubmitted 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   1
         Top             =   4560
         Width           =   5412
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "YYYY/MM/DD Format"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   3360
         TabIndex        =   23
         Top             =   5120
         Width           =   1932
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Mailed at"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   120
         TabIndex        =   21
         Top             =   5120
         Width           =   1332
      End
      Begin VB.Label cCode 
         Caption         =   "cCode"
         Height          =   372
         Left            =   5520
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   120
         TabIndex        =   13
         Top             =   320
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   120
         TabIndex        =   12
         Top             =   800
         Width           =   1212
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   120
         TabIndex        =   11
         Top             =   1280
         Width           =   1212
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1212
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "URL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   1212
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Compatibility"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   120
         TabIndex        =   8
         Top             =   4160
         Width           =   1212
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Submitted on"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   120
         TabIndex        =   7
         Top             =   4640
         Width           =   1212
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIndex As Integer
Private Sub ClearFields()

   On Error Resume Next
   Call CallStackPush("frmMain", "ClearFields")
   
   cCode.Caption = ""
   cTitle.Text = ""
   cCategory.Text = ""
   cLevel.Text = ""
   cDes.Text = ""
   cURL.url = ""
   cCompatibility.Text = ""
   cSubmitted.Text = ""
   cMailDate.Text = ""
   cReviewed.Value = 0

   Call CallStackPop

End Sub

Sub DeleteRecord()

   Dim SQL As String
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMain", "DeleteRecord")
   
   Ans = MsgBox("Do you realy want to delete this record?", vbYesNo)
   
   If Ans = 6 Then
      SQL = "DELETE * From Code WHERE (((Code.Code)=" & cCode.Caption & "));"
      gCurrentDB.Execute SQL
   
      Call RefreshTreeView
      Call ClearFields
      Call EnableFields(False)
      Call EnableButtons(False)
   End If

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "DeleteRecord", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub

Sub EnableButtons(Status As Boolean)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMain", "EnableButtons")

   Toolbar1.Buttons(1).Enabled = Status
   Toolbar1.Buttons(2).Enabled = Status
   Toolbar1.Buttons(4).Enabled = Status
   Toolbar1.Buttons(6).Enabled = Status

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMain", "EnableButtons", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select
   
End Sub

Sub EnableFields(Status As Boolean)

   On Error Resume Next
   Call CallStackPush("frmMain", "EnableFields")
   
   cCode.Enabled = Status
   cTitle.Enabled = Status
   cCategory.Enabled = Status
   cLevel.Enabled = Status
   cDes.Enabled = Status
   cURL.Enabled = Status
   cCompatibility.Enabled = Status
   cSubmitted.Enabled = Status
   cMailDate.Enabled = Status

   Call CallStackPop

End Sub


Sub ExitSub()

   On Error Resume Next

   Unload Me

End Sub

Sub FillBookmarks()

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMain", "FillBookmarks")
   
   SQL = "SELECT Code.Code, Code.Title, Code.Level, Code.Category, Code.Bookmark"
   SQL = SQL + " From Code"
   SQL = SQL + " WHERE (((Code.Bookmark)=True))"
   SQL = SQL + " ORDER BY Code.Category;"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   If GetNumbRecsRs(Rs) > 0 Then
      Rs.MoveFirst: Row% = 1
      Do Until Rs.EOF
         Key$ = "Code_" & Rs(0)
         cBookmarksList.ListItems.Add , Key$, Rs(1)
         cBookmarksList.ListItems(Row%).SubItems(1) = Rs(2)
         cBookmarksList.ListItems(Row%).SubItems(2) = Rs(3)
         Rs.MoveNext: Row% = Row% + 1
      Loop
   End If
   
   Rs.Close
   Set Rs = Nothing

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "FillBookmarks", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub

Sub InsertRecord()

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMain", "InsertRecord")
   
   SQL = "SELECT Code.Code, Code.Title, Code.Category, Code.Level, Code.Description, Code.URL, Code.Compatibility, Code.Submitted, Code.MailDate"
   SQL = SQL + " From Code"
   SQL = SQL + " WHERE (((Code.Code)=" & cCode.Caption & "));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenDynaset)
   
   If GetNumbRecsRs(Rs) > 0 Then
      Rs.Edit
      Rs(0) = cCode.Caption
      Rs(1) = cTitle.Text
      Rs(2) = cCategory.Text
      Rs(3) = cLevel.Text
      Rs(4) = cDes.Text
      Rs(5) = cURL.url
      Rs(6) = cCompatibility.Text
      Rs(7) = cSubmitted.Text
      Rs(8) = cMailDate.Text
      Rs.Update
   End If
   
   Rs.Close
   Set Rs = Nothing
   
   Call RefreshTreeView
   Call ClearFields
   Call EnableFields(False)
   Call EnableButtons(False)

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "InsertRecord", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub

Private Sub BookmarkRecord()

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMain", "BookmarkRecord")
   
   If Len(cCode.Caption) = 0 Then Call CallStackPop: Exit Sub
   
   SQL = "SELECT Code.Bookmark From Code WHERE (((Code.Code)=" & cCode.Caption & "));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenDynaset)
   
   Rs.Edit
   Rs(0) = True
   Rs.Update
   
   Rs.Close
   Set Rs = Nothing
   
   Row% = cBookmarksList.ListItems.Count + 1
   
   Key$ = "Code_" & cCode.Caption
   cBookmarksList.ListItems.Add , Key$, cTitle.Text
   cBookmarksList.ListItems(Row%).SubItems(1) = cLevel.Text
   cBookmarksList.ListItems(Row%).SubItems(2) = cCategory.Text

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "BookmarkRecord", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub

Private Sub cBookmarksList_ItemClick(ByVal Item As MSComctlLib.ListItem)

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMain", "cBookmarksList_ItemClick")
   
   If Item.Key = "Root" Then Call CallStackPop: Exit Sub
   If Left$(Item.Key, 3) = "Cat" Then Call CallStackPop: Exit Sub
   
   Txt = Item.Key
   LIndex = Item.Index
   code = Right$(Txt, Len(Txt) - 5)
   
   SQL = "SELECT DISTINCTROW Code.Code, Code.Title, Code.Category, Code.Level, Code.Description, Code.URL, Code.Compatibility, Code.Submitted, Code.Reviewed"
   SQL = SQL + " From Code"
   SQL = SQL + " WHERE (((Code.Code)=" & code & "));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   Call ClearFields
   If GetNumbRecsRs(Rs) > 0 Then
      cCode.Caption = Rs(0)
      cTitle.Text = Rs(1)
      cCategory.Text = Rs(2)
      cLevel.Text = Rs(3)
      cDes.Text = Rs(4)
      cURL.url = Rs(5)
      cCompatibility.Text = Rs(6)
      cSubmitted.Text = Rs(7)
      If Rs(8) Then cReviewed.Value = 1
      If Not Rs(8) Then cReviewed.Value = 0
      Call EnableFields(True)
   End If
   
   Rs.Close
   Set Rs = Nothing

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "TreeView1_NodeClick", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub


Private Sub cBookmarksList_KeyUp(KeyCode As Integer, Shift As Integer)

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMain", "cBookmarksList_KeyUp")
   
   If KeyCode = vbKeyDelete Then
      Txt = cBookmarksList.ListItems(LIndex).Key
      code = Right$(Txt, Len(Txt) - 5)
      SQL = "SELECT Code.Bookmark From Code WHERE (((Code.Code)=" & cCode.Caption & "));"
      Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenDynaset)
   
      Rs.Edit
      Rs(0) = False
      Rs.Update
   
      Rs.Close
      Set Rs = Nothing
      
      cBookmarksList.ListItems.Remove LIndex
      If cBookmarksList.ListItems.Count > 0 Then cBookmarksList.ListItems(1).Selected = True
      Call ClearFields
      TreeView1.SetFocus
      TreeView1.Nodes("Root").EnsureVisible
      TreeView1.Nodes("Root").Selected = True
   End If

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "cBookmarkList_KeyUp", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub


Private Sub cURL_AfterClick(ByVal Failed As Boolean)

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMain", "cURL_AfterClick")

   cReviewed.Value = 1

   SQL = "SELECT Code.Reviewed From Code WHERE (((Code.Code)=" & cCode.Caption & "));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenDynaset)
   
   Rs.Edit
   Rs(0) = True
   Rs.Update
   
   Rs.Close
   Set Rs = Nothing
   
   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "cURL_AfterClick", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub

Private Sub cURL_BeforeClick(Cancel As Boolean)

   On Error Resume Next
   Call CallStackPush("frmMain", "cURL_BeforeClick")
   
   If cReviewed.Value = 1 Then
      Ans = MsgBox("You have already visit this web address. Do you want to see it again?", vbYesNo)
      If Ans = 7 Then Cancel = True
   End If

   Call CallStackPop

End Sub

Private Sub Form_Load()

   On Error GoTo ErrorHandler
   Call CallStackPush("frmMain", "Form_Load")
   
   frmSplash.Show
   frmSplash!cActions.Caption = "Retrieving options"
   frmSplash.Refresh
   
   gstDBFolder = GetSetting("PSCDatabase", "Initialize", "Database_Folder")
   gstDBName = GetSetting("PSCDatabase", "Initialize", "Database_Name")
   gstPOP3 = GetSetting("PSCDatabase", "Initialize", "POP3_Mail_Server")
   gstSMTP = GetSetting("PSCDatabase", "Initialize", "SMTP_Mail_Server")
   gstLoginName = GetSetting("PSCDatabase", "Initialize", "Login_Name")
   gstPassword = GetSetting("PSCDatabase", "Initialize", "Password")
   gstMailAddress = GetSetting("PSCDatabase", "Initialize", "Email_Address")
   gstYourName = GetSetting("PSCDatabase", "Initialize", "Your_Name")
   gstMessageRules = GetSetting("PSCDatabase", "Initialize", "Message_Rules")

   If Len(gstDBFolder) > 0 And Len(gstDBName) > 0 Then
      DBFileName$ = Dir$(gstDBFolder + gstDBName)
      If Len(DBFileName$) > 0 Then
         Set gCurrentDB = OpenDatabase(App.Path + "\PSCDatabase.mdb")
      Else
         MsgBox "The program can't find the database at " & gstDBFolder & gstDBName, vbCritical, "Planet Source Code Database"
         Call CallStackPop
         Unload Me
      End If
   Else
      frmPreferences.Show 1, Me
   End If
   
   Set gCurrentDB = OpenDatabase(gstDBFolder + gstDBName)
   
   Call RefreshTreeView
   Call ClearFields
   Call EnableFields(False)
   Call FillBookmarks
   
   cBookmarksList.ColumnHeaders(1).Width = 6240
   cBookmarksList.ColumnHeaders(2).Width = 1440
   cBookmarksList.ColumnHeaders(3).Width = 2800
   
   Unload frmSplash

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "Form_Load", Err.Number, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub


Private Sub Form_Unload(Cancel As Integer)

   On Error Resume Next
   
   gCurrentDB.Close
   Set gCurrentDB = Nothing
   
   End

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

   On Error Resume Next
   Call CallStackPush("frmMain", "Toolbar1_ButtonClick")

   Select Case Button.Key
      Case "cInsert"
         Call InsertRecord
      Case "cDelete"
         Call DeleteRecord
      Case "cImport"
         frmImport.Show
      Case "cClear"
         Call ClearFields
         Call EnableFields(False)
         Call EnableButtons(False)
      Case "cSearch"
         frmSearch.Show
      Case "cBookMark"
         Call BookmarkRecord
      Case "cMail"
         frmMail.Show
      Case "cRecent"
         frmRecent.Show
      Case "cPref"
         frmPreferences.Show
      Case "cOldMail"
         frmViewMails.Show
      Case "cAbout"
         frmSplash.Show
         frmSplash!cActions.Visible = False
      Case "cExit"
         Call ExitSub
   End Select

   Call CallStackPop

End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmImport", "TreeView1_NodeClick")
   
   If Node.Key = "Root" Then Call CallStackPop: Exit Sub
   If Left$(Node.Key, 3) = "Cat" Then Call CallStackPop: Exit Sub
   
   Txt = Node.Key
   code = Right$(Txt, Len(Txt) - 5)
   
   SQL = "SELECT DISTINCTROW Code.Code, Code.Title, Code.Category, Code.Level, Code.Description, Code.URL, Code.Compatibility, Code.Submitted, Code.Reviewed, Code.MailDate"
   SQL = SQL + " From Code"
   SQL = SQL + " WHERE (((Code.Code)=" & code & "));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   Call ClearFields
   If GetNumbRecsRs(Rs) > 0 Then
      cCode.Caption = Rs(0)
      cTitle.Text = Rs(1)
      cCategory.Text = Rs(2)
      cLevel.Text = Rs(3)
      cDes.Text = Rs(4)
      cURL.url = Rs(5)
      cCompatibility.Text = Rs(6)
      cSubmitted.Text = Rs(7)
      If Rs(8) Then cReviewed.Value = 1
      If Not Rs(8) Then cReviewed.Value = 0
      cMailDate.Text = Rs(9)
      Call EnableFields(True)
      Call EnableButtons(True)
   End If
   
   Rs.Close
   Set Rs = Nothing

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "TreeView1_NodeClick", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub
