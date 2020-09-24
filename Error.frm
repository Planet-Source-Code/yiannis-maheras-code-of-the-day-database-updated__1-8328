VERSION 5.00
Begin VB.Form fError 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ËÜèïò ÐñïãñÜììáôïò"
   ClientHeight    =   3768
   ClientLeft      =   3336
   ClientTop       =   2676
   ClientWidth     =   7356
   HelpContextID   =   1070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3768
   ScaleWidth      =   7356
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cMailSend 
      Caption         =   "Send Mail"
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
      Left            =   5052
      Picture         =   "Error.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton cIgnore 
      Caption         =   "Continue"
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
      Left            =   3852
      Picture         =   "Error.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   852
   End
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
      Height          =   852
      Left            =   1332
      Picture         =   "Error.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   852
   End
   Begin VB.CommandButton cRetry 
      Caption         =   "Retry"
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
      Left            =   2532
      Picture         =   "Error.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   972
   End
   Begin VB.TextBox cErrLine 
      Alignment       =   1  'Right Justify
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
      Left            =   5760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   1452
   End
   Begin VB.TextBox cErrDes 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   7095
   End
   Begin VB.TextBox cErrNo 
      Alignment       =   1  'Right Justify
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
      Left            =   5760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1452
   End
   Begin VB.TextBox cRoutine 
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
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   2892
   End
   Begin VB.TextBox cFormName 
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
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label cStatus 
      BackStyle       =   0  'Transparent
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
      Left            =   0
      TabIndex        =   14
      Top             =   3480
      Width           =   7332
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Line Number"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Error description"
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
      TabIndex        =   7
      Top             =   1080
      Width           =   7092
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Error code"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Routine"
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
      TabIndex        =   5
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Form"
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
      TabIndex        =   4
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "fError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
Option Explicit
Option Compare Text
Dim bAuthLogin      As Boolean
Dim MyEncodeType    As ENCODE_METHOD

Private Sub cCancel_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("fError", "cCancel_Click")

   gstErrorFlag = "CANCEL"
   Unload Me

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("fError", "cCancel_Click", Err, Erl())
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


Private Sub cIgnore_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("fError", "cIgnore_Click")

   gstErrorFlag = "IGNORE"
   Unload Me

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("fError", "cIgnore_Click", Err, Erl())
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

Private Sub cMailSend_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("fError", "cMailSend_Click")

   Dim Subject As String
   Dim body As String
   Dim starSeparator As String
   Dim i
   
   Subject = "Error in Planet Source Code Database"

   ' write error message into the text file
   starSeparator = String(70, "*")
   body = starSeparator & vbCrLf

   ' error source procedure name
   body = body + "* Source: " & cFormName.Text & "(" & cRoutine.Text & ")" & vbCrLf

   ' define procedure section containing the error
   body = body + "* Error Number: " & cErrNo.Text & vbCrLf
   body = body + "* Error Line Number: " & cErrLine.Text & vbCrLf
   body = body + "* Description:" & vbCrLf

   ' save sErrorDescription string in predefined format
   body = body + cErrDes.Text & vbCrLf

   ' call cascade description
   body = body + "* Error Call History: " & vbCrLf
   For i = 0 To CallStackSize - 1
      body = body + "*    - " & CallStack(i) & vbCrLf
   Next i

   ' put the time stamp
   body = body + "* Date/Time: " & Now & vbCrLf
   body = body + starSeparator
   
   With poSendMail
      .SMTPHost = gstSMTP
      .From = gstMailAddress
      .FromDisplayName = gstYourName
      .Recipient = "salemlot@otenet.gr"
      .RecipientDisplayName = "Yiannis Maheras"
      .ReplyToAddress = gstMailAddress
      .Subject = Subject
      .Message = body
      .Send
   End With
   
   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("fError", "cMailSend_Click", Err, Erl())
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

Private Sub cRetry_Click()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("fError", "cRetry_Click")

   gstErrorFlag = "RETRY"
   Unload Me

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("fError", "cRetry_Click", Err, Erl())
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
   Call CallStackPush("fError", "Form_Load")

   Set poSendMail = New clsSendMail

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("fError", "Form_Load", Err, Erl())
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


Private Sub poSendMail_SendFailed(Explanation As String)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("fError", "poSendMail_SendFailed")

   MsgBox ("Your attempt to send mail failed for the following reason: " & vbCrLf & Explanation)

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("fError", "poSendMail_SendFailed", Err, Erl())
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

Private Sub poSendMail_SendSuccesful()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("fError", "poSendMail_SendSuccesful")

   MsgBox "Mail send", vbOKOnly

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("fError", "poSendMail_SendSuccesful", Err, Erl())
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

Private Sub poSendMail_Status(Status As String)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("fError", "poSendMail_Status")

   cStatus.Caption = Status

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("fError", "poSendMail_Status", Err, Erl())
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



