VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "CCRPFTV6.OCX"
Begin VB.Form frmImport 
   Caption         =   "Importing Code into the Database"
   ClientHeight    =   5436
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10176
   LinkTopic       =   "Form1"
   ScaleHeight     =   5436
   ScaleWidth      =   10176
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView cFiles 
      Height          =   4212
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   6012
      _ExtentX        =   10605
      _ExtentY        =   7430
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin CCRPFolderTV6.FolderTreeview cFolder 
      Height          =   4152
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   7324
   End
   Begin VB.CommandButton cSellectAll 
      Caption         =   "Sellect All"
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
      Left            =   4560
      Picture         =   "frmImport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   972
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
      Left            =   6600
      Picture         =   "frmImport.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
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
      Left            =   2640
      Picture         =   "frmImport.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   972
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FO_DELETE = &H3
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Dim ImportFlag As Boolean
Dim FilesDates()
Private Sub cExit_Click()

   On Error Resume Next
   Call CallStackPush("frmImport", "cExit_Click")
   
   If ImportFlag Then Call RefreshTreeView
   Unload Me

   Call CallStackPop

End Sub

Private Sub cFolder_FolderClick(Folder As CCRPFolderTV6.Folder, Location As CCRPFolderTV6.ftvHitTestConstants)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmImport", "cFolder_FolderClick")

   Dim lngHandle As Long, SHDirOp As SHFILEOPSTRUCT, lngLong As Long
   Dim Ft1 As FILETIME, Ft2 As FILETIME, SysTime As SYSTEMTIME
    
   Me.MousePointer = 11
   
   MyPath = cFolder.SelectedFolder
   If Right$(MyPath, 1) <> "\" Then MyPath = MyPath + "\"
   MyPath = MyPath + "*.txt"
   MyName = Dir(MyPath): Row% = 0
   Do While MyName <> ""
      MyName = Dir: Row% = Row% + 1
   Loop

   If Row% = 0 Then
      Me.MousePointer = 0
      Call CallStackPop
      Exit Sub
   End If
   ReDim FilesDates(1 To Row%, 1 To 2)

   cFiles.ListItems.Clear
   MyPath = cFolder.SelectedFolder
   If Right$(MyPath, 1) <> "\" Then MyPath = MyPath + "\"
   MyPath = MyPath + "*.txt"
   MyName = Dir(MyPath): Row% = 1
   Do While MyName <> ""
      FilesDates(Row%, 1) = MyName
      
      TempPath = cFolder.SelectedFolder
      If Right$(TempPath, 1) <> "\" Then TempPath = TempPath + "\"
      
      'Open the file
      lngHandle = CreateFile(TempPath + MyName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
      
      'Get the fil's time
      GetFileTime lngHandle, Ft1, Ft1, Ft2
      
      'Convert the file time to the local file time
      FileTimeToLocalFileTime Ft2, Ft1
      
      'Convert the file time to system file time
      FileTimeToSystemTime Ft1, SysTime
      If Len(Trim(Str$(SysTime.wMonth))) = 1 Then
         m$ = "0" + Trim(Str$(SysTime.wMonth))
      Else
         m$ = Trim(Str$(SysTime.wMonth))
      End If
      If Len(Trim(Str$(SysTime.wDay))) = 1 Then
         d$ = "0" + Trim(Str$(SysTime.wDay))
      Else
         d$ = Trim(Str$(SysTime.wDay))
      End If
      FileDate = LTrim(Str$(SysTime.wYear)) + "/" + m$ + "/" + d$
      FilesDates(Row%, 2) = FileDate
      
      'Close the file
      CloseHandle lngHandle
      
      MyName = Dir: Row% = Row% + 1
   Loop
   
   Call TableSort(1, UBound(FilesDates))
   For i% = 1 To UBound(FilesDates)
      cFiles.ListItems.Add , , FilesDates(i%, 1)
      cFiles.ListItems(i%).SubItems(1) = FilesDates(i%, 2)
   Next i%
   
   cFiles.Refresh
   
   Me.MousePointer = 0

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmImport", "cFolder_FolderClick", Err, Erl())
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
Sub TableSort(lLbound As Long, lUbound As Long)

   '::: Passed:strListString array:::'
   '::: lLboundLower bound to sort (usually 1) :::'
   '::: lUboundUpper bound to sort (usually ubound()) :::'
    
   Dim strTemp As String
   Dim strBuffer As String
   Dim lngCurLow As Long
   Dim lngCurHigh As Long
   Dim lngCurMidpoint As Long
   
   On Error Resume Next
   Call CallStackPush("frmImport", "TableSort")
   
   lngCurLow = lLbound ' Start current low and high at actual low/high
   lngCurHigh = lUbound
    
   If lUbound <= lLbound Then Call CallStackPop: Exit Sub     ' Error!
   lngCurMidpoint = (lLbound + lUbound) \ 2 ' Find the approx midpoint of the array
    
   strTemp = FilesDates(lngCurMidpoint, 2) ' Pick as a starting point (we are making
   ' an assumption that the data *might* be
   ' in semi-sorted order already!

   Do While (lngCurLow <= lngCurHigh)
      Do While FilesDates(lngCurLow, 2) < strTemp
         lngCurLow = lngCurLow + 1
         If lngCurLow = lUbound Then Exit Do
      Loop
      Do While strTemp < FilesDates(lngCurHigh, 2)
         lngCurHigh = lngCurHigh - 1
         If lngCurHigh = lLbound Then Exit Do
      Loop
      If (lngCurLow <= lngCurHigh) Then ' if low is <= high then swap
         GoSub Swap
         lngCurLow = lngCurLow + 1 ' CurLow++
         lngCurHigh = lngCurHigh - 1 ' CurLow--
      End If
   Loop

   If lLbound < lngCurHigh Then ' Recurse if necessary
      TableSort lLbound, lngCurHigh
   End If

   If lngCurLow < lUbound Then ' Recurse if necessary
      TableSort lngCurLow, lUbound
   End If

   Call CallStackPop

Exit Sub

Swap:
   Dim TempTable(1 To 1, 1 To 2)
   
   For i% = 1 To 2
      TempTable(1, i%) = FilesDates(lngCurLow, i%)
      FilesDates(lngCurLow, i%) = ""
      FilesDates(lngCurLow, i%) = FilesDates(lngCurHigh, i%)
      FilesDates(lngCurHigh, i%) = TempTable(1, i%)
   Next i%
   
   Erase TempTable
   Return

   Call CallStackPop
End Sub

Private Sub cImport_Click()

   On Error GoTo ErrorHandler
   Call CallStackPush("frmImport", "cImport_Click")
   
   Me.MousePointer = 11
   
   Counter% = 0
   Proccess% = 0
   For i% = 1 To cFiles.ListItems.Count
      If cFiles.ListItems(i%).Checked Then Counter% = Counter% + 1
   Next i%

   If Counter% = 0 Then Call CallStackPop: Exit Sub
   ImportFlag = True

   Load frmImportControls
   frmProgress.Show
   frmProgress!cAni.Play
   For j% = 1 To cFiles.ListItems.Count
      If cFiles.ListItems(j%).Checked Then
         FileName = cFolder.SelectedFolder + "\" + cFiles.ListItems(j%).Text
         frmProgress!cProgress.Value = (Proccess% / Counter%) * 100
         frmProgress!cFile.Caption = FileName
         frmProgress!cMailsImport.Caption = "Mails Imported: " & Proccess% & " from " & Counter%
         frmProgress.Refresh
      
         fl% = FreeFile
         Open FileName For Input As fl%
   
         Row% = 1
         Do Until EOF(fl%)
            Line Input #fl%, A$
            If Len(Trim(A$)) > 0 Then
               frmImportControls!cMessageText.TextMatrix(Row%, 0) = A$
               frmImportControls!cMessageText.rows = frmImportControls!cMessageText.rows + 1
               Row% = Row% + 1
            End If
         Loop
   
         Close #fl%
   
         Call ImportMails

         If UCase(frmImportControls!cCodeExist.Caption) = "CANCEL" Then Exit For
         Proccess% = Proccess% + 1
      End If
   Next j%
   Unload frmImportControls
   Unload frmProgress
   
   Me.MousePointer = 0

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg("frmImport", "cImport_Click", Err, Erl())
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

Private Sub cSellectAll_Click()

   On Error Resume Next
   Call CallStackPush("frmImport", "cSellectAll_Click")
   
   For i% = 0 To cFiles.ListItems.Count
      cFiles.ListItems(i%).Checked = True
   Next i%

   Call CallStackPop

End Sub


Private Sub Form_Load()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmImport", "Form_Load")

   cFiles.ColumnHeaders(1).Width = 4630
   cFiles.ColumnHeaders(2).Width = 1070

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmImport", "Form_Load", Err, Erl())
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



