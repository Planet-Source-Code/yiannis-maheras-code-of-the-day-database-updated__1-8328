VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{57A90BB1-9F57-11D3-B479-A0A072A969C6}#7.0#0"; "VBWHYPERLINK.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Search form"
   ClientHeight    =   8064
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7212
   LinkTopic       =   "Form1"
   ScaleHeight     =   8064
   ScaleWidth      =   7212
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cClear 
      Caption         =   "Clear"
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
      Left            =   4920
      TabIndex        =   22
      Top             =   600
      Width           =   972
   End
   Begin MSComctlLib.ListView cResults 
      Height          =   1452
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   2561
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
      TabIndex        =   7
      Top             =   2880
      Width           =   6972
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
         Height          =   372
         Left            =   1440
         TabIndex        =   27
         Top             =   4560
         Width           =   1812
      End
      Begin VB.CommandButton cBookmark 
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
         Height          =   732
         Left            =   5760
         Picture         =   "frmSearch.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3240
         Width           =   1092
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
         Height          =   372
         Left            =   1440
         TabIndex        =   13
         Top             =   4080
         Width           =   5412
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
         Height          =   372
         Left            =   1440
         TabIndex        =   12
         Top             =   3600
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
         TabIndex        =   11
         Top             =   1680
         Width           =   5412
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
         Height          =   372
         Left            =   1440
         TabIndex        =   10
         Top             =   1200
         Width           =   3612
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
         Height          =   372
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   4332
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
         Height          =   372
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   5412
      End
      Begin Hyperlink.vbwHyperlink cURL 
         Height          =   348
         Left            =   1440
         TabIndex        =   23
         Top             =   3120
         Width           =   3612
         _ExtentX        =   6371
         _ExtentY        =   614
         HoverColour     =   16711680
         Caption         =   "Planet Source Code"
         AutoSize        =   0   'False
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
      Begin VB.Label Label10 
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
         TabIndex        =   28
         Top             =   4640
         Width           =   2292
      End
      Begin VB.Label Label9 
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
         TabIndex        =   26
         Top             =   4640
         Width           =   1332
      End
      Begin VB.Label cCode 
         Caption         =   "Label9"
         Height          =   372
         Left            =   5280
         TabIndex        =   25
         Top             =   1200
         Visible         =   0   'False
         Width           =   1572
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
         TabIndex        =   20
         Top             =   4160
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
         TabIndex        =   19
         Top             =   3680
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
         Height          =   372
         Left            =   120
         TabIndex        =   18
         Top             =   3120
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
         TabIndex        =   17
         Top             =   1680
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
         TabIndex        =   16
         Top             =   1280
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
         TabIndex        =   15
         Top             =   800
         Width           =   1212
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
         TabIndex        =   14
         Top             =   320
         Width           =   1212
      End
   End
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
      Height          =   372
      Left            =   6120
      TabIndex        =   6
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton cSearch 
      Caption         =   "Go!"
      Default         =   -1  'True
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
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   972
   End
   Begin VB.OptionButton cDesSearch 
      Caption         =   "Search in description"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   2052
   End
   Begin VB.OptionButton cTitleSearch 
      Caption         =   "Search in Title"
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
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   2052
   End
   Begin VB.TextBox cSearchString 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   5772
   End
   Begin VB.Label cResultsLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   6972
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   7200
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   7200
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search string"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1212
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cBookMark_Click()

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmSearch", "cBookmark_Click")
   
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


Private Sub cClear_Click()

   On Error Resume Next
   Call CallStackPush("frmSearch", "cClear_Click")
   
   cSearchString.Text = ""
   cTitleSearch.Value = True
   cResults.ListItems.Clear
   Call ClearFields

   Call CallStackPop

End Sub


Private Sub cExit_Click()

   On Error Resume Next
   
   Unload Me

End Sub



Private Sub ClearFields()

   On Error Resume Next
   Call CallStackPush("frmSearch", "ClearFields")
   
   cCode.Caption = ""
   cTitle.Text = ""
   cCategory.Text = ""
   cLevel.Text = ""
   cDes.Text = ""
   cURL.URL = ""
   cCompatibility.Text = ""
   cSubmitted.Text = ""

   Call CallStackPop

End Sub


Private Sub cResults_ItemClick(ByVal Item As MSComctlLib.ListItem)

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmSearch", "cResults_ItemClick")
   
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
   Call ErrMsg(Me.Name, "cResults_ItemClick", Err, Erl())
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


Private Sub cSearch_Click()

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmSearch", "cSearch_Click")
   
   Me.MousePointer = 11
   
   SQL = "SELECT DISTINCTROW Code.Title, Code.Category, Code.Code From Code"
   If cTitleSearch.Value Then SQL = SQL + " WHERE (((Code.Title) Like '*" & cSearchString.Text & "*'));"
   If cDesSearch.Value Then SQL = SQL + " WHERE (((Code.Description) Like '*" & cSearchString.Text & "*'));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   If GetNumbRecsRs(Rs) > 0 Then
      cResultsLabel.Caption = "found " & GetNumbRecsRs(Rs) & " records"
      Row% = 1
      Do Until Rs.EOF
         Key$ = "Code_" & Rs(2)
         cResults.ListItems.Add , Key$, Rs(0)
         cResults.ListItems(Row%).SubItems(1) = Rs(1)
         Rs.MoveNext: Row% = Row% + 1
      Loop
   Else
      cResultsLabel.Caption = "No matches found"
   End If
   
   Rs.Close
   Set Rs = Nothing
   
   Me.MousePointer = 0
   
   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.Name, "cSearch_Click", Err, Erl())
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
   Call CallStackPush("frmSearch", "cURL_AfterClick")

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

   On Error Resume Next
   Call CallStackPush("frmSearch", "cURL_BeforeClick")
   
   If cReviewed.Value = 1 Then
      Ans = MsgBox("You have already visit this web address. Do you want to see it again?", vbYesNo)
      If Ans = 7 Then Cancel = True
   End If

   Call CallStackPop

End Sub

Private Sub Form_Load()

   On Error Resume Next
   Call CallStackPush("frmSearch", "Form_Load")
   
   cResults.ColumnHeaders(1).Width = 4430
   cResults.ColumnHeaders(2).Width = 2230

   Call CallStackPop

End Sub
