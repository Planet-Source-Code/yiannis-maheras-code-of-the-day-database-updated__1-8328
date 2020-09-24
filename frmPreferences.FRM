VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPreferences 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   5172
   ClientLeft      =   1104
   ClientTop       =   1848
   ClientWidth     =   9144
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   17
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5172
   ScaleWidth      =   9144
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cRules 
      Caption         =   "E-mail Rules"
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
      Left            =   4746
      Picture         =   "frmPreferences.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4080
      Width           =   972
   End
   Begin VB.TextBox cYourName 
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
      Left            =   2040
      TabIndex        =   18
      Top             =   3480
      Width           =   3852
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
      Height          =   972
      Left            =   6066
      Picture         =   "frmPreferences.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   972
   End
   Begin VB.CommandButton cFind 
      Caption         =   "Find DB"
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
      Left            =   3426
      Picture         =   "frmPreferences.frx":25F4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   972
   End
   Begin VB.CommandButton cInsert 
      Caption         =   "Insert"
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
      Left            =   2106
      Picture         =   "frmPreferences.frx":2EBE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   972
   End
   Begin VB.TextBox cPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2520
      Width           =   2172
   End
   Begin VB.TextBox cLoginName 
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
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   2172
   End
   Begin VB.TextBox cSMTPMailServer 
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
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   3852
   End
   Begin VB.TextBox cPOP3MailServer 
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   3852
   End
   Begin VB.TextBox cMailAddress 
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
      Left            =   2040
      TabIndex        =   6
      Top             =   3000
      Width           =   3852
   End
   Begin VB.TextBox cDBName 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   3852
   End
   Begin VB.TextBox cDBFolder 
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
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   6972
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   600
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name"
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
      TabIndex        =   17
      Top             =   3480
      Width           =   1932
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      TabIndex        =   16
      Top             =   2520
      Width           =   3252
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
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
      TabIndex        =   15
      Top             =   2040
      Width           =   3252
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your SMTP Mail Server"
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
      TabIndex        =   14
      Top             =   1560
      Width           =   3252
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Your POP3 Mail Server"
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
      TabIndex        =   13
      Top             =   1080
      Width           =   3732
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Your E-Mail Address"
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
      TabIndex        =   12
      Top             =   3000
      Width           =   3252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Databse Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   3012
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Database Folder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3012
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cExit_Click()
   
   On Error Resume Next

   Unload Me

End Sub


Private Sub cFind_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmPreferences", "cFind_Click")

   CommonDialog1.Filter = "Database (*.mdb)|*.mdb"
   CommonDialog1.InitDir = "c:\"
   CommonDialog1.ShowOpen
   
   Txt = Trim$(CommonDialog1.FileName)
   If Right$(Txt, 3) = "mdb" Then
      Do
         A% = InStr(Txt, "\")
         If A% = 0 Then
            cDBName.Text = Txt
            A% = InStr(CommonDialog1.FileName, Txt)
            cDBFolder.Text = Left$(CommonDialog1.FileName, A% - 1)
            Exit Do
         End If
         Txt = Right$(Txt, Len(Txt) - A%)
      Loop
   End If
   
   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmPreferences", "cFind_Click", Err, Erl())
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


Private Sub cInsert_Click()

   Dim Txt As String

   On Error GoTo ErrorHandler
   Call CallStackPush("frmPreferences", "cInsert_Click")

   Screen.MousePointer = 11
   
   If gstDBFolder <> cDBFolder.Text Or gstDBName <> cDBName.Text Then
      gstDBFolder = cDBFolder.Text
      gstDBName = cDBName.Text

      DB$ = cDBFolder.Text + cDBName.Text
      fi$ = Dir$(DB$)
      If Len(fi$) > 0 Then
         gstPOP3 = cPOP3MailServer.Text
         gstSMTP = cSMTPMailServer.Text
         gstLoginName = cLoginName.Text
         gstPassword = Encode(cPassword.Text)
         gstMailAddress = cMailAddress.Text
         gstYourName = cYourName.Text
         SaveSetting "PSCDatabase", "Initialize", "Database_Folder", gstDBFolder
         SaveSetting "PSCDatabase", "Initialize", "Database_Name", gstDBName
         SaveSetting "PSCDatabase", "Initialize", "POP3_Mail_Server", gstPOP3
         SaveSetting "PSCDatabase", "Initialize", "SMTP_Mail_Server", gstSMTP
         SaveSetting "PSCDatabase", "Initialize", "Login_Name", gstLoginName
         SaveSetting "PSCDatabase", "Initialize", "Password", gstPassword
         SaveSetting "PSCDatabase", "Initialize", "Email_Address", gstMailAddress
         SaveSetting "PSCDatabase", "Initialize", "Your_Name", gstYourName
         Set gCurrentDB = OpenDatabase(DB$)
      Else
         MsgBox "Database could not be found", 0
         SaveSetting "PSCDatabase", "Initialize", "Database_Folder", gstDBFolder
         SaveSetting "PSCDatabase", "Initialize", "Database_Name", gstDBName
         gstDBFolder = cDBFolder.Text
         gstDBName = cDBName.Text
      End If
   Else
      gstPOP3 = cPOP3MailServer.Text
      gstSMTP = cSMTPMailServer.Text
      gstLoginName = cLoginName.Text
      gstPassword = Encode(cPassword.Text)
      gstMailAddress = cMailAddress.Text
      gstYourName = cYourName.Text
      SaveSetting "PSCDatabase", "Initialize", "POP3_Mail_Server", gstPOP3
      SaveSetting "PSCDatabase", "Initialize", "SMTP_Mail_Server", gstSMTP
      SaveSetting "PSCDatabase", "Initialize", "Login_Name", gstLoginName
      SaveSetting "PSCDatabase", "Initialize", "Password", gstPassword
      SaveSetting "PSCDatabase", "Initialize", "Email_Address", gstMailAddress
      SaveSetting "PSCDatabase", "Initialize", "Your_Name", gstYourName
   End If
   
   Screen.MousePointer = 0
   
   Unload Me

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg(Me.name, "InsertData", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         cDBFolder.Text = GetSetting("PSCDatabase", "Initialize", "Database_Folder")
         cDBName.Text = GetSetting("PSCDatabase", "Initialize", "Database_Name")
         gstDBFolder = cDBFolder.Text
         gstDBName = cDBName.Text
         Screen.MousePointer = 0
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub

Private Sub cRules_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmPreferences", "cRules_Click")

   frmMailRules.Show

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmPreferences", "cRules_Click", Err, Erl())
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
   Call CallStackPush("frmPreferences", "Form_Load")

   Me.MousePointer = 11

   cDBFolder.Text = gstDBFolder
   cDBName.Text = gstDBName
   cPOP3MailServer.Text = gstPOP3
   cSMTPMailServer.Text = gstSMTP
   cLoginName.Text = gstLoginName
   cPassword.Text = DeCode(gstPassword)
   cMailAddress.Text = gstMailAddress
   cYourName.Text = gstYourName

   Me.MousePointer = 0

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg("frmPreferences", "Form_Load", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Screen.MousePointer = 0
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub
Private Sub Form_Unload(Cancel As Integer)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmPreferences", "Form_Unload")

   Set fPreferences = Nothing

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmPreferences", "Form_Unload", Err, Erl())
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
