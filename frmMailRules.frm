VERSION 5.00
Begin VB.Form frmMailRules 
   Caption         =   "E-mail Rules"
   ClientHeight    =   4704
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6036
   LinkTopic       =   "Form1"
   ScaleHeight     =   4704
   ScaleWidth      =   6036
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cCancel 
      Caption         =   "Cancel"
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
      Left            =   3480
      Picture         =   "frmMailRules.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   972
   End
   Begin VB.CommandButton cOk 
      Caption         =   "OK"
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
      Left            =   1560
      Picture         =   "frmMailRules.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   972
   End
   Begin VB.CommandButton cRemoveRule 
      Caption         =   "Remove Rule"
      Enabled         =   0   'False
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
      Left            =   4560
      TabIndex        =   4
      Top             =   1200
      Width           =   1332
   End
   Begin VB.CommandButton cAddRule 
      Caption         =   "Add New Rule"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   720
      Width           =   1332
   End
   Begin VB.ListBox cRulesList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2208
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4212
   End
   Begin VB.TextBox cNewRule 
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
      TabIndex        =   1
      Top             =   720
      Width           =   4212
   End
   Begin VB.Label cListindex 
      Caption         =   "-1"
      Height          =   252
      Left            =   4560
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Here you can type the message rules for Planet Source Code, in order the E-mail client to receive only this messages"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5772
   End
End
Attribute VB_Name = "frmMailRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cAddRule_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMailRules", "cAddRule_Click")

   If cListindex.Caption = "-1" Then
      cRulesList.AddItem cNewRule.Text
   Else
      cRulesList.List(cListindex.Caption) = cNewRule.Text
      cListindex.Caption = "-1"
   End If

   cNewRule.Text = ""
   cNewRule.SetFocus
   
   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMailRules", "cAddRule_Click", Err, Erl())
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


Private Sub cCancel_Click()
   
   On Error Resume Next

   Unload Me

End Sub


Private Sub cOk_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMailRules", "cOk_Click")

   gstMessageRules = ""
   For i% = 0 To cRulesList.ListCount - 1
      gstMessageRules = gstMessageRules + cRulesList.List(i%) + "|"
   Next i%
   gstMessageRules = Left$(gstMessageRules, Len(gstMessageRules) - 1)
   
   SaveSetting "PSCDatabase", "Initialize", "Message_Rules", gstMessageRules
   
   Call CallStackPop
   Unload Me

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMailRules", "cOk_Click", Err, Erl())
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

Private Sub cRemoveRule_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMailRules", "cRemoveRule_Click")

   lS% = cRulesList.ListIndex
   cRulesList.RemoveItem lS%
   cListindex.Caption = "-1"
   cNewRule.Text = ""
   
   cRemoveRule.Enabled = False
   cNewRule.SetFocus
   
   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMailRules", "cRemoveRule_Click", Err, Erl())
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

Private Sub cRulesList_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmMailRules", "cRulesList_Click")

   If cRulesList.ListCount > 0 Then
      cRemoveRule.Enabled = True
      cNewRule.Text = cRulesList.List(cRulesList.ListIndex)
      cListindex.Caption = cRulesList.ListIndex
   End If

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMailRules", "cRulesList_Click", Err, Erl())
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
   Call CallStackPush("frmMailRules", "Form_Load")

   TempString = gstMessageRules
   Do
      A% = InStr(TempString, "|")
      If A% > 0 Then
         cRulesList.AddItem Left$(TempString, A% - 1)
         TempString = Right$(TempString, Len(TempString) - A%)
      Else
         cRulesList.AddItem TempString
         Exit Do
      End If
   Loop

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmMailRules", "Form_Load", Err, Erl())
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
