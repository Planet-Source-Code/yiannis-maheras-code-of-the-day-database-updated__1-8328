VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{57A90BB1-9F57-11D3-B479-A0A072A969C6}#7.0#0"; "VBWHYPERLINK.OCX"
Begin VB.Form frmViewMails 
   Caption         =   "View older Mails"
   ClientHeight    =   8184
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10452
   LinkTopic       =   "Form1"
   ScaleHeight     =   8184
   ScaleWidth      =   10452
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7212
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   3132
      _ExtentX        =   5525
      _ExtentY        =   12721
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   4080
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewMails.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewMails.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewMails.frx":11B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewMails.frx":1A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewMails.frx":2970
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewMails.frx":324C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   684
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10452
      _ExtentX        =   18436
      _ExtentY        =   1207
      ButtonWidth     =   1461
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Bookmark"
            Key             =   "cBookmark"
            Object.ToolTipText     =   "Bookmark this code"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Clear"
            Key             =   "cClear"
            Object.ToolTipText     =   "Clears the contents of the fields"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "New Mail"
            Key             =   "cNewMail"
            Object.ToolTipText     =   "Select an other mail"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Export"
            Key             =   "cExport"
            Object.ToolTipText     =   "Export this e-mail in a text file"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete"
            Key             =   "cDelete"
            Object.ToolTipText     =   "Delete the current mail"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "cExit"
            Object.ToolTipText     =   "Exit this form"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   5052
      Left            =   3360
      TabIndex        =   1
      Top             =   3000
      Width           =   6972
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
         TabIndex        =   10
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
         Height          =   360
         Left            =   1440
         TabIndex        =   9
         Top             =   3600
         Width           =   5412
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
         TabIndex        =   8
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
         Height          =   360
         Left            =   1440
         TabIndex        =   7
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
         Height          =   360
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   3612
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
         TabIndex        =   5
         Top             =   240
         Width           =   5412
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
         TabIndex        =   4
         Top             =   4560
         Width           =   1572
      End
      Begin VB.CheckBox cReviewed 
         Caption         =   "Shows if this code has been reviewed"
         Height          =   852
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   1212
      End
      Begin Hyperlink.vbwHyperlink cURL 
         Height          =   228
         Left            =   1440
         TabIndex        =   3
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
      Begin VB.Label cCode 
         Caption         =   "Label9"
         Height          =   372
         Left            =   5280
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label10 
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
         TabIndex        =   19
         Top             =   4160
         Width           =   1212
      End
      Begin VB.Label Label8 
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
         TabIndex        =   18
         Top             =   3680
         Width           =   1212
      End
      Begin VB.Label Label7 
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
         Height          =   204
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   1212
      End
      Begin VB.Label Label6 
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
         TabIndex        =   16
         Top             =   1680
         Width           =   1212
      End
      Begin VB.Label Label5 
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
         TabIndex        =   15
         Top             =   1280
         Width           =   1212
      End
      Begin VB.Label Label4 
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
         TabIndex        =   14
         Top             =   800
         Width           =   1212
      End
      Begin VB.Label Label3 
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
         TabIndex        =   12
         Top             =   4640
         Width           =   1332
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
         Height          =   200
         Left            =   3120
         TabIndex        =   11
         Top             =   4640
         Width           =   2172
      End
   End
   Begin MSComctlLib.ListView cRecentMails 
      Height          =   2172
      Left            =   3360
      TabIndex        =   0
      Top             =   840
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   3831
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
End
Attribute VB_Name = "frmViewMails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DeleteFlag As Boolean
Sub BookmarkCode()

   Dim SQL As String
   Dim Rs As Recordset
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmViewMails", "BookmarkCode")
   
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
   Call ErrMsg(Me.Name, "BookmarkCode", Err, Erl())
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

Private Sub ClearContents()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmViewMails", "ClearContents")

   Toolbar1.Buttons(1).Enabled = False
   Toolbar1.Buttons(2).Enabled = False
   Call ClearFields

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmViewMails", "ClearContents", Err, Erl())
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

Sub DeleteMail()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmViewMails", "DeleteMail")

   Dim SQL As String

   Code = TreeView1.SelectedItem.Key
      
   Ans% = MsgBox("Do you realy want to delete this mail from the Database?", vbYesNo)
   If Ans% = 6 Then
      Me.MousePointer = 11

      SQL = "DELETE * From Code Where (((Code.MailDate) = " & Code & "));"
      gCurrentDB.Execute SQL
      DeleteFlag = True
      Call NewMail

      Me.MousePointer = 0
   End If
  
   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmViewMails", "DeleteMail", Err, Erl())
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

Private Sub Export2Txt()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmViewMails", "Export2Txt")

   Dim SQL As String
   Dim Rs As Recordset

   Code = TreeView1.SelectedItem.Key
   
   fl% = FreeFile
   FileName$ = App.Path + "\Planet Source Code Mail - " & Code & ".txt"
   
   Do
      A% = InStr(FileName$, "/")
      If A% = 0 Then Exit Do
      Mid$(FileName$, A%) = "-"
   Loop
   Open FileName$ For Output As fl%

   SQL = "SELECT * From Code Where (((Code.MailDate) = '" & Code & "')) ORDER BY Code.Code;"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)

   GoSub FindDay
   Print #fl%, wDay$ & ", " & Rs(10)
   Print #fl%, String(48, "*")
   Print #fl%, "Table of contents"
   
   Row% = 1
   Do Until Rs.EOF
      Print #fl%, Row% & ")" & Rs(1)
      Rs.MoveNext: Row% = Row% + 1
   Loop
 
   Print #fl%, String(48, "*")
   
   Rs.MoveFirst: Row% = 1
   Do Until Rs.EOF
      Print #fl%, Row% & ")" & Rs(1)
      Print #fl%, ""
      Print #fl%, "Category: " & Rs(2)
      Print #fl%, "Level: " & Rs(3)
      Print #fl%, ""
      GoSub ParseDesc
      Print #fl%, ""
      Print #fl%, "Complete source code is at:"
      Print #fl%, Rs(5)
      Print #fl%, ""
      Print #fl%, "Compatibility: " & Rs(6)
      Print #fl%, "Submitted on " & Rs(7)
      Print #fl%, String(48, "*")
      Rs.MoveNext: Row% = Row% + 1
   Loop
   
   Rs.Close
   Set Rs = Nothing
   
   Close #fl%

   Call CallStackPop

Exit Sub

FindDay:
   Select Case Weekday(CDate(Rs(10)))
      Case 1
         wDay$ = "Sunday"
      Case 2
         wDay$ = "Monday"
      Case 3
         wDay$ = "Tuesday"
      Case 4
         wDay$ = "Wednesday"
      Case 5
         wDay$ = "Thursday"
      Case 6
         wDay$ = "Friday"
      Case 7
         wDay$ = "Saturday"
      Case Else
   End Select
   Return
   
ParseDesc:
   FirstTime = True
   DescString = Rs(4)
   Do
      If FirstTime Then
         No% = 63
      Else
         No% = 76
      End If
      If Len(DescString) <= No% Then
         If FirstTime Then
            Print #fl%, "Description: " & DescString
         Else
            Print #fl%, DescString
         End If
         Exit Do
      Else
         If Mid$(DescString, No%, 1) <> " " Then
            For i% = No% To 1 Step -1
               If Mid$(DescString, i%, 1) = " " Then Exit For
            Next i%
            If FirstTime Then
               Print #fl%, "Description: " & Left$(DescString, i% - 1)
               DescString = Trim(Right$(DescString, Len(DescString) - i%))
            Else
               Print #fl%, Left$(DescString, i% - 1)
               DescString = Trim(Right$(DescString, Len(DescString) - i%))
            End If
         Else
            If FirstTime Then
               Print #fl%, "Description: " & Left$(DescString, No% - 1)
               DescString = Trim(Right$(DescString, Len(DescString) - No%))
            Else
               Print #fl%, Left$(DescString, No% - 1)
               DescString = Trim(Right$(DescString, Len(DescString) - No%))
            End If
         End If
      End If
      FirstTime = False
   Loop
   Return
   
Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmViewMails", "Export2Txt", Err, Erl())
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

Sub FillDatesTree()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmViewMails", "FillDatesTree")

   Dim SQL As String
   Dim YearRs As Recordset
   Dim MonthRs As Recordset
   Dim DatesRs As Recordset
   
   SQL = "SELECT Year([MailDate]) AS Expr1 From Code GROUP BY Year([MailDate]);"
   Set YearRs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   Set nodX = TreeView1.Nodes.Add(, , "Root", "Old E-Mails")
   
   If GetNumbRecsRs(YearRs) > 0 Then
      Do Until YearRs.EOF
         Key$ = "Year_" & YearRs(0)
         Set nodX = TreeView1.Nodes.Add("Root", tvwChild, Key$, "Mail from year " & YearRs(0))
         
         SQL = "SELECT First(Code.MailDate) AS FirstOfMailDate, Month([MailDate]) AS Expr1 From Code"
         SQL = SQL + " GROUP BY Month([MailDate])"
         SQL = SQL + " HAVING (((First(Code.MailDate)) Like '" & YearRs(0) & "*'));"
         Set MonthRs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)

         If GetNumbRecsRs(MonthRs) > 0 Then
            Do Until MonthRs.EOF
               MonthKey$ = "Year_" & YearRs(0) & "_Month_" & MonthRs(1)
               GoSub FindMonthName
               Set nodX = TreeView1.Nodes.Add(Key$, tvwChild, MonthKey$, MName$)
               
               If Len(MonthRs(1)) = 1 Then
                  m$ = "0" + CStr(MonthRs(1))
               Else
                  m$ = MonthRs(1)
               End If
               SQL = "SELECT Code.MailDate From Code GROUP BY Code.MailDate Having (((Code.MailDate) Like '" & YearRs(0) & "/" & m$ & "/" & "*')) ORDER BY Code.MailDate;"
               Set DatesRs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)

               If GetNumbRecsRs(DatesRs) > 0 Then
                  Do Until DatesRs.EOF
                     DateKey$ = DatesRs(0)
                     Set nodX = TreeView1.Nodes.Add(MonthKey$, tvwChild, DateKey$, DatesRs(0))
                     DatesRs.MoveNext
                  Loop
               End If
               
               DatesRs.Close
               Set DatesRs = Nothing
               
               MonthRs.MoveNext
            Loop
         End If
         
         MonthRs.Close
         Set MonthRs = Nothing
         
         YearRs.MoveNext
      Loop
      TreeView1.Nodes(Key$).EnsureVisible
   End If
   
   YearRs.Close
   Set YearRs = Nothing

   Call CallStackPop

Exit Sub

FindMonthName:
   Select Case MonthRs(1)
      Case 1
         MName$ = "January"
      Case 2
         MName$ = "Febraury"
      Case 3
         MName$ = "March"
      Case 4
         MName$ = "April"
      Case 5
         MName$ = "May"
      Case 6
         MName$ = "June"
      Case 7
         MName$ = "July"
      Case 8
         MName$ = "Agoust"
      Case 9
         MName$ = "September"
      Case 10
         MName$ = "October"
      Case 11
         MName$ = "November"
      Case 12
         MName$ = "December"
      Case Else
   End Select
   Return
Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmViewMails", "FillDatesTree", Err, Erl())
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

Private Sub NewMail()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmViewMails", "NewMail")

   cRecentMails.ListItems.Clear
   Call ClearFields

   Toolbar1.Buttons(1).Enabled = False
   Toolbar1.Buttons(2).Enabled = False
   Toolbar1.Buttons(3).Enabled = False
   Toolbar1.Buttons(4).Enabled = False
   Toolbar1.Buttons(5).Enabled = False

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmViewMails", "NewMail", Err, Erl())
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

Private Sub cRecentMails_ItemClick(ByVal Item As MSComctlLib.ListItem)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmViewMails", "cRecentMails_ItemClick")

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

      Toolbar1.Buttons(1).Enabled = True
      Toolbar1.Buttons(2).Enabled = True
   End If

   Rs.Close
   Set Rs = Nothing

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmViewMails", "cRecentMails_ItemClick", Err, Erl())
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
   Call CallStackPush("frmViewMails", "ClearFields")

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
   Call ErrMsg("frmViewMails", "ClearFields", Err, Erl())
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
   Call CallStackPush("frmViewMails", "Form_Load")

   Call FillDatesTree
   cRecentMails.ColumnHeaders(1).Width = 4410
   cRecentMails.ColumnHeaders(2).Width = 2270

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmViewMails", "Form_Load", Err, Erl())
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
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmViewMails", "Toolbar1_ButtonClick")

   Select Case Button.Key
      Case "cBookmark"
         Call BookmarkCode
      Case "cClear"
         Call ClearContents
      Case "cNewMail"
         Call NewMail
      Case "cExport"
         Call Export2Txt
      Case "cDelete"
         Call DeleteMail
      Case "cExit"
         If DeleteFlag Then Call RefreshTreeView
         Unload Me
      Case Else
   End Select

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmViewMails", "Toolbar1_ButtonClick", Err, Erl())
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


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmViewMails", "TreeView1_NodeClick")

   Dim SQL As String
   Dim Rs As Recordset

   If Node.Key = "Root" Then Call CallStackPop: Exit Sub
   If Left$(Node.Key, 4) = "Year" Then Call CallStackPop: Exit Sub
   
   Code = Node.Key
      
   SQL = "SELECT Code.Code, Code.Title, Code.Category, Code.MailDate From Code Where (((Code.MailDate) = '" & Code & "')) ORDER BY Code.Code;"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)

   If GetNumbRecsRs(Rs) > 0 Then
      Row% = 1
      Do Until Rs.EOF
         Key$ = "Code_" & Rs(0)
         cRecentMails.ListItems.Add , Key$, Rs(1)
         cRecentMails.ListItems(Row%).SubItems(1) = Rs(2)
         Rs.MoveNext: Row% = Row% + 1
      Loop
      Toolbar1.Buttons(3).Enabled = True
      Toolbar1.Buttons(4).Enabled = True
      Toolbar1.Buttons(5).Enabled = True
   End If
   
   Rs.Close
   Set Rs = Nothing

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmViewMails", "TreeView1_NodeClick", Err, Erl())
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
