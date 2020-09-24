VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importing Procedure Progress"
   ClientHeight    =   2040
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   8616
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   8616
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.Animation cAni 
      Height          =   1452
      Left            =   7080
      TabIndex        =   3
      Top             =   120
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   2561
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   121
      FullHeight      =   121
   End
   Begin MSComctlLib.ProgressBar cProgress 
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   6852
      _ExtentX        =   12086
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label cRecsImport 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Records Imported: "
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
      Left            =   3720
      TabIndex        =   5
      Top             =   1680
      Width           =   3252
   End
   Begin VB.Label cMailsImport 
      BackStyle       =   0  'Transparent
      Caption         =   "Mails Imported: "
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
      TabIndex        =   4
      Top             =   1680
      Width           =   3492
   End
   Begin VB.Label cFile 
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
      Height          =   732
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6852
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proccessing File:"
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
      Width           =   2292
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   On Error Resume Next
   Call CallStackPush("frmProgress", "Form_Load")

   cAni.AutoPlay = True
   cAni.Open App.Path + "\store.avi"

   Call CallStackPop

End Sub

