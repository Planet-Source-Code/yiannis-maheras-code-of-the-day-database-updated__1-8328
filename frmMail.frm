VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B09269F3-5C29-11D3-A358-08002B000001}#1.0#0"; "FPOP301.OCX"
Begin VB.Form frmMail 
   Caption         =   "Receive PSC Mail's"
   ClientHeight    =   7512
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11304
   LinkTopic       =   "Form1"
   ScaleHeight     =   7512
   ScaleWidth      =   11304
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cOutlook 
      Caption         =   "Outlook"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   6006
      Picture         =   "frmMail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   972
   End
   Begin FreePOPcontrol.FreePOP FreePOP1 
      Left            =   8520
      Top             =   6360
      _ExtentX        =   1080
      _ExtentY        =   868
      POPHostname     =   ""
      POPPort         =   110
      POPUsername     =   ""
      POPPassword     =   ""
      POPTimeout      =   10
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
      Height          =   852
      Left            =   7686
      Picture         =   "frmMail.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   972
   End
   Begin VB.CommandButton cImport 
      Caption         =   "Import"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   4326
      Picture         =   "frmMail.frx":25F4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   972
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   7140
      Width           =   11304
      _ExtentX        =   19939
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19445
            MinWidth        =   8819
         EndProperty
      EndProperty
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
   Begin VB.CommandButton cReceive 
      Caption         =   "Receive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   2646
      Picture         =   "frmMail.frx":2EBE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   972
   End
   Begin VB.TextBox cMailBody 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3612
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2520
      Width           =   11000
   End
   Begin MSComctlLib.ListView cMailsList 
      Height          =   2292
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11000
      _ExtentX        =   19410
      _ExtentY        =   4043
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Received"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid cMessageText 
      Height          =   6012
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   11052
      _ExtentX        =   19495
      _ExtentY        =   10605
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mails()

Private Sub cExit_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMail", "cExit_Click")

   Call RefreshTreeView
   Unload Me

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMail", "cExit_Click", Err, Erl())
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

Private Sub cImport_Click()

   On Error GoTo ErrorHandler
   Call CallStackPush("frmMail", "cImport_Click")
   
   Me.MousePointer = 11
   
   Counter% = 0
   Proccess% = 0
   For i% = 1 To cMailsList.ListItems.Count
      If cMailsList.ListItems(i%).Checked Then Counter% = Counter% + 1
   Next i%

   If Counter% = 0 Then Call CallStackPop: Exit Sub

   Load frmImportControls
   frmProgress.Show
   frmProgress!cAni.Play
   For i% = 1 To cMailsList.ListItems.Count
      If cMailsList.ListItems(i%).Checked = True Then
         frmProgress!cProgress.Value = (Proccess% / Counter%) * 100
         frmProgress!cFile.Caption = Mails(i%, 3)
         frmProgress!cMailsImport.Caption = "Mails Imported: " & Proccess% & " from " & Counter%
         frmProgress.Refresh
         
         TempString = Mails(i%, 5)
   
         Row% = 1
         Do
            A% = InStr(TempString, Chr$(13))
            b% = InStr(TempString, Chr$(10))
            If A% > 0 And b% > 0 Then
               frmImportControls!cMessageText.TextMatrix(Row%, 0) = Left$(TempString, A% - 1)
               TempString = Right$(TempString, Len(TempString) - b%)
               Row% = Row% + 1
               frmImportControls!cMessageText.rows = frmImportControls!cMessageText.rows + 1
            Else
               Exit Do
            End If
         Loop
   
         Row% = 1
         Do
            If frmImportControls!cMessageText.rows - 1 < Row% Then Exit Do
            If Len(frmImportControls!cMessageText.TextMatrix(Row%, 0)) = 0 Then
               frmImportControls!cMessageText.RemoveItem Row%
            Else
               Row% = Row% + 1
            End If
         Loop
   
         Call ImportMails
         
         Proccess% = Proccess% + 1
      End If
   Next i%
   Unload frmImportControls
   Unload frmProgress
   
   Me.MousePointer = 0

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "cImport_Click", Err, Erl())
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

Private Sub cMailsList_ItemClick(ByVal Item As MSComctlLib.ListItem)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMail", "cMailsList_ItemClick")

   code = Right$(Item.Key, Len(Item.Key) - 5)
   cMailBody.Text = Mails(code, 5)

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMail", "cMailsList_ItemClick", Err, Erl())
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


Private Sub cOutlook_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMail", "cOutlook_Click")

   frmOutlook.Show

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMail", "cOutlook_Click", Err, Erl())
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

Private Sub cReceive_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMail", "cReceive_Click")

   Dim intCount As Integer

   cReceive.Enabled = False
   
   FreePOP1.POPHostname = gstPOP3
   FreePOP1.POPUsername = gstLoginName
   FreePOP1.POPPassword = DeCode(gstPassword)
   FreePOP1.Connect
   
   If FreePOP1.Error = 0 Then
      Row% = 1
      ReDim Mails(1 To FreePOP1.MsgCount, 1 To 5)
      For intCount = 1 To FreePOP1.MsgCount
         StatusBar1.Panels(1).Text = "Receiving message " & intCount & " from " & FreePOP1.MsgCount
         FreePOP1.SetCurrentMsg (intCount)
         A% = InStr(gstMessageRules, FreePOP1.MsgFrom)
         If A% > 0 Then
            If FreePOP1.Error = 0 Then
               FreePOP1.RetrieveCurrentMsg
               Key$ = "Mail_" & Row%
               cMailsList.ListItems.Add , Key$, FreePOP1.MsgFrom
               cMailsList.ListItems(Row%).SubItems(1) = FreePOP1.MsgSubject
               cMailsList.ListItems(Row%).SubItems(2) = FreePOP1.MsgDate
               cMailsList.ListItems(Row%).Checked = True
               Mails(Row%, 1) = Row%
               Mails(Row%, 2) = FreePOP1.MsgFrom
               Mails(Row%, 3) = FreePOP1.MsgSubject
               Mails(Row%, 4) = FreePOP1.MsgDate
               Mails(Row%, 5) = FreePOP1.MsgContents
               DoEvents
               If FreePOP1.Error <> 0 Then
                  MsgBox FreePOP1.ErrorText, vbOKOnly, "Error during RetrieveCurrentMsg"
               End If
            Else
               MsgBox FreePOP1.ErrorText, vbOKOnly, "Error during SetCurrentMsg"
            End If
            Row% = Row% + 1
         End If
      Next
      FreePOP1.Disconnect
   Else
      MsgBox FreePOP1.ErrorText, vbOKOnly, "Error during Connect"
   End If

   cReceive.Enabled = True

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMail", "cReceive_Click", Err, Erl())
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
   Call CallStackPush("frmMail", "Form_Load")

   cMailsList.ColumnHeaders(1).Width = 3000
   cMailsList.ColumnHeaders(2).Width = 5650
   cMailsList.ColumnHeaders(3).Width = 2270

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMail", "Form_Load", Err, Erl())
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
