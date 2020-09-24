VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{57A90BB1-9F57-11D3-B479-A0A072A969C6}#7.0#0"; "VBWHYPERLINK.OCX"
Begin VB.Form frmRecent 
   Caption         =   "Recent DB Additions"
   ClientHeight    =   8052
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10452
   LinkTopic       =   "Form1"
   ScaleHeight     =   8052
   ScaleWidth      =   10452
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   9360
      Picture         =   "frmRecent.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   972
   End
   Begin VB.CommandButton cBookMark 
      Caption         =   "Bookmark"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   9360
      Picture         =   "frmRecent.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   972
   End
   Begin MSComctlLib.ListView cRecentMails 
      Height          =   2772
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   10212
      _ExtentX        =   18013
      _ExtentY        =   4890
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5052
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   9132
      Begin VB.CheckBox cReviewed 
         Caption         =   "Shows if this code has been reviewed"
         Height          =   852
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1212
      End
      Begin Hyperlink.vbwHyperlink cURL 
         Height          =   228
         Left            =   1440
         TabIndex        =   20
         Top             =   3120
         Width           =   1644
         _ExtentX        =   2900
         _ExtentY        =   402
         HoverColour     =   16711680
         Caption         =   "Planet Source Code"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
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
         TabIndex        =   15
         Top             =   4560
         Width           =   1812
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
         Width           =   7572
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
         Height          =   1332
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1680
         Width           =   7572
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
         Top             =   3600
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
         Top             =   4080
         Width           =   5412
      End
      Begin VB.Label cCode 
         Caption         =   "Label9"
         Height          =   372
         Left            =   5160
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label9 
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
         Height          =   204
         Left            =   3360
         TabIndex        =   19
         Top             =   4644
         Width           =   2172
      End
      Begin VB.Label Label1 
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
         TabIndex        =   14
         Top             =   4640
         Width           =   1332
      End
      Begin VB.Label Label2 
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
      Begin VB.Label Label3 
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
      Begin VB.Label Label4 
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
      Begin VB.Label Label5 
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
         Height          =   200
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1212
      End
      Begin VB.Label Label6 
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
         Height          =   200
         Left            =   120
         TabIndex        =   9
         Top             =   3200
         Width           =   1212
      End
      Begin VB.Label Label7 
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
         Top             =   3680
         Width           =   1212
      End
      Begin VB.Label Label8 
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
         Top             =   4160
         Width           =   1212
      End
   End
End
Attribute VB_Name = "frmRecent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cBookMark_Click()

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmRecent", "cBookMark_Click")
   
   If Len(cCode.Caption) = 0 Then Call CallStackPop: Exit Sub
   
   SQL = "SELECT Code.Bookmark From Code WHERE (((Code.Code)=" & cCode.Caption & "));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenDynaset)
   
   Rs.Edit
   Rs(0) = True
   Rs.Update
   
   Rs.Close
   Set Rs = Nothing
   
   Row% = frmMain!cBookmarksList.ListItems.Count + 1
   
   Key$ = "Code_" & cCode.Caption
   frmMain!cBookmarksList.ListItems.Add , Key$, cTitle.Text
   frmMain!cBookmarksList.ListItems(Row%).SubItems(1) = cLevel.Text
   frmMain!cBookmarksList.ListItems(Row%).SubItems(2) = cCategory.Text

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.Name, "cBookmark_Click", Err, Erl())
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

Private Sub cExit_Click()
   
   On Error Resume Next

   Unload Me

End Sub

Private Sub cRecentMails_ItemClick(ByVal Item As MSComctlLib.ListItem)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmRecent", "cRecentMails_ItemClick")

   Dim SQL As String
   Dim Rs As Recordset
   
   CodeKey = Item.Key

   Code = Right$(CodeKey, Len(CodeKey) - 5)
   
   SQL = "SELECT DISTINCTROW Code.Code, Code.Title, Code.Category, Code.Level, Code.Description, Code.URL, Code.Compatibility, Code.Submitted, Code.MailDate"
   SQL = SQL + " From Code"
   SQL = SQL + " WHERE (((Code.Code)=" & Code & "));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   Call ClearFields
   If GetNumbRecsRs(Rs) > 0 Then
      cCode.Caption = Rs(0)
      cTitle.Text = Rs(1)
      cCategory.Text = Rs(2)
      cLevel.Text = Rs(3)
      cDes.Text = Rs(4)
      cURL.URL = Rs(5)
      cCompatibility.Text = Rs(6)
      cSubmitted.Text = Rs(7)
      cMailDate.Text = Rs(8)
   End If
   
   Rs.Close
   Set Rs = Nothing

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmRecent", "cRecentMails_ItemClick", Err, Erl())
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


Private Sub ClearFields()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmRecent", "ClearFields")

   cTitle.Text = ""
   cCategory.Text = ""
   cLevel.Text = ""
   cDes.Text = ""
   cURL.URL = ""
   cCompatibility.Text = ""
   cSubmitted.Text = ""
   cMailDate.Text = ""

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmRecent", "ClearFields", Err, Erl())
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
   Call CallStackPush("frmRecent", "cURL_AfterClick")

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
   Call ErrMsg(Me.Name, "cURL_AfterClick", Err, Erl())
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
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmRecent", "cURL_BeforeClick")

   If cReviewed.Value = 1 Then
      Ans = MsgBox("You have already visit this web address. Do you want to see it again?", vbYesNo)
      If Ans = 7 Then Cancel = True
   End If

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmRecent", "cURL_BeforeClick", Err, Erl())
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

Private Sub Form_Load()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmRecent", "Form_Load")

   Dim SQL As String
   Dim Rs As Recordset

   SQL = "SELECT TOP 1 Code.Code, Code.MailDate From Code ORDER BY Code.MailDate DESC;"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   If GetNumbRecsRs(Rs) > 0 Then LastMail = Rs(1)
   
   Rs.Close
   Set Rs = Nothing
   
   If Len(LastMail) > 0 Then
      SQL = "SELECT Code.Code, Code.Title, Code.Category From Code Where (((Code.MailDate) = '" & LastMail & "')) ORDER BY Code.Code;"
      Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)

      If GetNumbRecsRs(Rs) > 0 Then
         Row% = 1
         Do Until Rs.EOF
            Key$ = "Code_" & Rs(0)
            cRecentMails.ListItems.Add , Key$, Rs(1)
            cRecentMails.ListItems(Row%).SubItems(1) = Rs(2)
            Rs.MoveNext: Row% = Row% + 1
         Loop
      End If
   
      Rs.Close
      Set Rs = Nothing
      
      Me.Caption = Me.Caption + " - Last addition in Databse at " & LastMail & " (YYYY/MM/DD format)"
      cRecentMails.ListItems(1).Selected = True
   Else
         Call CallStackPop
      Call cExit_Click
   End If
   
   cRecentMails.ColumnHeaders(1).Width = 6410
   cRecentMails.ColumnHeaders(2).Width = 3510

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmRecent", "Form_Load", Err, Erl())
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
