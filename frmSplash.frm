VERSION 5.00
Object = "{57A90BB1-9F57-11D3-B479-A0A072A969C6}#7.0#0"; "VBWHYPERLINK.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7224
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7224
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hyperlink.vbwHyperlink vbwHyperlink1 
      Height          =   228
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   1656
      _ExtentX        =   2921
      _ExtentY        =   402
      HoverColour     =   16711680
      Caption         =   "salemlot@otenet.gr"
      URL             =   "salemlot@otenet.gr"
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
   Begin VB.Label cActions 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   6972
   End
   Begin VB.Label cVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version: "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   3720
      TabIndex        =   0
      Top             =   4080
      Width           =   2052
   End
   Begin VB.Image Image2 
      Height          =   5544
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   
   On Error Resume Next

   If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("frmSplash", "Form_Load")

   cVersion.Caption = "Version " & App.Major & " . " & App.Minor & " . " & App.Revision
   Me.Width = Image2.Width
   Me.Height = Image2.Height

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("frmSplash", "Form_Load", Err, Erl())
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
