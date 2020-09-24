VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutlook 
   Caption         =   "Outlook Client"
   ClientHeight    =   8052
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10212
   LinkTopic       =   "Form1"
   ScaleHeight     =   8052
   ScaleWidth      =   10212
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cSelectAll 
      Caption         =   "Select All"
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
      Left            =   5280
      Picture         =   "frmOutlook.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   972
   End
   Begin VB.CommandButton cCancel 
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
      Left            =   6600
      Picture         =   "frmOutlook.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit the form"
      Top             =   6960
      Width           =   972
   End
   Begin VB.CommandButton cExport 
      Caption         =   "Export"
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
      Left            =   3960
      Picture         =   "frmOutlook.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Export the selected e-mails in text files"
      Top             =   6960
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
      Height          =   972
      Left            =   2640
      Picture         =   "frmOutlook.frx":205E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Import the selected e-mail into the database"
      Top             =   6960
      Width           =   972
   End
   Begin VB.TextBox cBody 
      Height          =   3852
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmOutlook.frx":2928
      Top             =   3000
      Width           =   6852
   End
   Begin MSComctlLib.ListView cMails 
      Height          =   2772
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   6852
      _ExtentX        =   12086
      _ExtentY        =   4890
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sender"
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
   Begin MSComctlLib.TreeView treOutl 
      Height          =   6732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3012
      _ExtentX        =   5313
      _ExtentY        =   11875
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim these objects in module level declarations code
Private objFolder As MapiFolder

Private Sub cCancel_Click()
   
   On Error Resume Next

   Call RefreshTreeView
   Unload Me

End Sub

Private Sub cExport_Click()

   On Error GoTo ErrorHandler
   Call CallStackPush("frmOutlook", "cExport_Click")
   
   Me.MousePointer = 11
   
   Counter% = 0
   Proccess% = 0
   For i% = 1 To cMails.ListItems.Count
      If cMails.ListItems(i%).Checked Then Counter% = Counter% + 1
   Next i%

   If Counter% = 0 Then Call CallStackPop: Exit Sub

   For i% = 1 To cMails.ListItems.Count
      If cMails.ListItems(i%).Checked = True Then
         With objFolder.Items(cMails.ListItems(i%).Index)
            fl% = FreeFile
            Open App.Path + "/Planet Source Code - " + Format$(.ReceivedTime, "yyyy-mm-dd") + ".txt" For Output As fl%
            Print #fl%, .body
         End With
      End If
   Next i%
   
   Me.MousePointer = 0

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "cExport_Click", Err, Erl())
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
   Call CallStackPush("frmOutlook", "cImport_Click")
   
   Me.MousePointer = 11
   
   Counter% = 0
   Proccess% = 0
   For i% = 1 To cMails.ListItems.Count
      If cMails.ListItems(i%).Checked Then Counter% = Counter% + 1
   Next i%

   If Counter% = 0 Then Call CallStackPop: Exit Sub

   Load frmImportControls
   frmProgress.Show
   frmProgress!cAni.Play
   For i% = 1 To cMails.ListItems.Count
      If cMails.ListItems(i%).Checked = True Then
         frmProgress!cProgress.Value = (Proccess% / Counter%) * 100
         frmProgress!cFile.Caption = objFolder.Items(cMails.ListItems(i%).Index).Subject
         frmProgress!cMailsImport.Caption = "Mails Imported: " & Proccess% & " from " & Counter%
         frmProgress.Refresh

         With objFolder.Items(cMails.ListItems(i%).Index)
            TempString = .body
         End With

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

Private Sub cMails_ItemClick(ByVal Item As MSComctlLib.ListItem)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmOutlook", "cMails_ItemClick")

   Me.MousePointer = 11
   
   With objFolder.Items(cMails.SelectedItem.Index)
      cBody.Text = .body
   End With

   Me.MousePointer = 0
   
   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmOutlook", "cMails_ItemClick", Err, Erl())
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


Private Sub cSelectAll_Click()

   On Error Resume Next
   Call CallStackPush("frmOutlook", "cSelectAll_Click")
   
   For i% = 0 To cMails.ListItems.Count
      cMails.ListItems(i%).Checked = True
   Next i%

   Call CallStackPop

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmOutlook", "Form_Unload")

   Set gobjOutlook = Nothing
   Set gobjNamespace = Nothing
   Set objFolder = Nothing

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmOutlook", "Form_Unload", Err, Erl())
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
   Call CallStackPush("frmOutlook", "Form_Load")

   Me.MousePointer = 11
   
   If Not CreateOutlookInstance Then
      MsgBox "Error", vbCritical
      End
   End If

   Filltree
   
   cMails.ColumnHeaders(1).Width = 3050
   cMails.ColumnHeaders(2).Width = 2150
   cMails.ColumnHeaders(3).Width = 1600
   
   Me.MousePointer = 0
    
   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmOutlook", "Form_Load", Err, Erl())
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


Sub Filltree()
   
   Dim colMapiFolders As Outlook.Folders
   Dim MapiFolder As Outlook.MapiFolder
   Dim colFolders As Outlook.Folders
   Dim Folder As Outlook.MapiFolder
   Dim nodX As Node
   Dim strMapiFolder As String
   Dim strFolder As String
   Dim strFolderKey As String
   Dim strKey As String

   On Error GoTo ErrorHandler
   Call CallStackPush("frmOutlook", "Filltree")

   'Begin the tree at the root
   Set nodX = treOutl.Nodes.Add(, , "root", "Outlook: " + gobjNamespace.CurrentUser)

   'Walk the MAPI Folder tree
   Set colMapiFolders = gobjNamespace.Folders

   For Each MapiFolder In colMapiFolders
      strMapiFolder = MapiFolder.name
      Set nodX = treOutl.Nodes.Add("root", tvwChild, strMapiFolder, strMapiFolder)
    
      'Folders within MapiFolders
      Set colFolders = MapiFolder.Folders

      For Each Folder In colFolders
         strFolder = Folder.name
         strFolderKey = Folder.EntryID
         Set nodX = treOutl.Nodes.Add(strMapiFolder, tvwChild, strFolderKey, strFolder)
         nodX.EnsureVisible
         
         'Folders within Folders
         EnumerateFolders Folder
      Next
   Next

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmOutlook", "Filltree", Err, Erl())
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

Function EnumerateFolders(objParentFolder As Object) As Integer

   Dim colchildFolders As Outlook.Folders
   Dim ChildFolder As Outlook.MapiFolder
   Dim strChildFolder As String
   Dim strChildFolderKey As String
   Dim strFolderKey As String
   Dim nodX As Node

   On Error GoTo ErrorHandler
   Call CallStackPush("frmOutlook", "EnumerateFolders")
   
   Set colchildFolders = objParentFolder.Folders

   If colchildFolders.Count <> 0 Then
      strFolderKey = objParentFolder.EntryID
      For Each ChildFolder In colchildFolders
         strChildFolder = ChildFolder.name
         strChildFolderKey = ChildFolder.EntryID
         Set nodX = treOutl.Nodes.Add(strFolderKey, tvwChild, strChildFolderKey, strChildFolder)
         EnumerateFolders ChildFolder
      Next
   End If

   Call CallStackPop

Exit Function
 
ErrorHandler:
   Call ErrMsg("frmOutlook", "EnumerateFolders", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Function
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select
   
End Function

Sub GetFolderContents()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmOutlook", "GetFolderContents")

   For i = 1 To objFolder.Items.Count
      With objFolder.Items(i)
         cMails.ListItems.Add , , .SenderName
         cMails.ListItems(i).SubItems(1) = .Subject
         cMails.ListItems(i).SubItems(2) = .ReceivedTime
      End With
   Next

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmOutlook", "GetFolderContents", Err, Erl())
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
Private Sub treOutl_NodeClick(ByVal Node As MSComctlLib.Node)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmOutlook", "treOutl_NodeClick")

   Me.MousePointer = 11
   
   Set objFolder = gobjNamespace.GetFolderFromID(Node.Key)
   GetFolderContents
   
   Me.MousePointer = 0

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmOutlook", "treOutl_NodeClick", Err, Erl())
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
