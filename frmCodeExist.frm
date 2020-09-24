VERSION 5.00
Begin VB.Form frmCodeExist 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6840
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   10932
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10932
   ShowInTaskbar   =   0   'False
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
      Height          =   852
      Left            =   6780
      Picture         =   "frmCodeExist.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
      Width           =   972
   End
   Begin VB.CommandButton cIgnore 
      Caption         =   "Ignore"
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
      Left            =   5580
      Picture         =   "frmCodeExist.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
      Width           =   972
   End
   Begin VB.CommandButton cUpdate 
      Caption         =   "Update"
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
      Left            =   4380
      Picture         =   "frmCodeExist.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5880
      Width           =   972
   End
   Begin VB.CommandButton cSave 
      Caption         =   "Save"
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
      Left            =   3180
      Picture         =   "frmCodeExist.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Height          =   5412
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   10692
      Begin VB.TextBox cMailDateNew 
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
         Left            =   6120
         TabIndex        =   33
         Top             =   4920
         Width           =   1812
      End
      Begin VB.TextBox cMailDateOld 
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
         TabIndex        =   31
         Top             =   4920
         Width           =   1812
      End
      Begin VB.TextBox cSubmittedNew 
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
         Left            =   6120
         TabIndex        =   26
         Top             =   4440
         Width           =   4452
      End
      Begin VB.TextBox cCompatibilityNew 
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
         Left            =   6120
         TabIndex        =   25
         Top             =   3960
         Width           =   4452
      End
      Begin VB.TextBox cURLNew 
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
         Left            =   6120
         TabIndex        =   24
         Top             =   3480
         Width           =   4452
      End
      Begin VB.TextBox cDesNew 
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
         Left            =   6120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   2040
         Width           =   4452
      End
      Begin VB.TextBox cLevelNew 
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
         Left            =   6120
         TabIndex        =   22
         Top             =   1560
         Width           =   3612
      End
      Begin VB.TextBox cCategoryNew 
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
         Left            =   6120
         TabIndex        =   21
         Top             =   1080
         Width           =   3612
      End
      Begin VB.TextBox cTitleNew 
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
         Left            =   6120
         TabIndex        =   20
         Top             =   600
         Width           =   4452
      End
      Begin VB.TextBox cTitleOld 
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
         TabIndex        =   8
         Top             =   600
         Width           =   4452
      End
      Begin VB.TextBox cCategoryOld 
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
         Top             =   1080
         Width           =   3612
      End
      Begin VB.TextBox cLevelOld 
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
         Top             =   1560
         Width           =   3612
      End
      Begin VB.TextBox cDesOld 
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
         TabIndex        =   5
         Top             =   2040
         Width           =   4452
      End
      Begin VB.TextBox cURLOld 
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
         Top             =   3480
         Width           =   4452
      End
      Begin VB.TextBox cCompatibilityOld 
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
         TabIndex        =   3
         Top             =   3960
         Width           =   4452
      End
      Begin VB.TextBox cSubmittedOld 
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
         TabIndex        =   2
         Top             =   4440
         Width           =   4452
      End
      Begin VB.Label Label14 
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
         Left            =   8040
         TabIndex        =   32
         Top             =   5000
         Width           =   1932
      End
      Begin VB.Label Label12 
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
         Height          =   204
         Left            =   3360
         TabIndex        =   30
         Top             =   5004
         Width           =   1932
      End
      Begin VB.Label Label11 
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
         TabIndex        =   29
         Top             =   5000
         Width           =   1332
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   6000
         X2              =   6000
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   5985
         X2              =   5985
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   6120
         TabIndex        =   28
         Top             =   240
         Width           =   4452
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Existing code in database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   4452
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
         TabIndex        =   15
         Top             =   680
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
         TabIndex        =   14
         Top             =   1160
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
         TabIndex        =   13
         Top             =   1640
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
         Height          =   372
         Left            =   120
         TabIndex        =   12
         Top             =   2040
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
         Height          =   200
         Left            =   120
         TabIndex        =   11
         Top             =   3560
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
         TabIndex        =   10
         Top             =   4040
         Width           =   1212
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
         TabIndex        =   9
         Top             =   4520
         Width           =   1212
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "It seems that this piece of code allready exist in the database. What you want me to do?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10692
   End
End
Attribute VB_Name = "frmCodeExist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cCancel_Click()

   On Error Resume Next
   Call CallStackPush("frmCodeExist", "cCancel_Click")
   
   frmImportControls!cCodeExist.Caption = "Cancel"
   Unload Me

   Call CallStackPop

End Sub

Private Sub cIgnore_Click()

   On Error Resume Next
   Call CallStackPush("frmCodeExist", "cIgnore_Click")
   
   frmImportControls!cCodeExist.Caption = "Ignore"
   Unload Me

   Call CallStackPop

End Sub

Private Sub cSave_Click()

   On Error Resume Next
   Call CallStackPush("frmCodeExist", "cSave_Click")
   
   frmImportControls!cCodeExist.Caption = "Save"
   Unload Me

   Call CallStackPop

End Sub

Private Sub cUpdate_Click()

   On Error Resume Next
   Call CallStackPush("frmCodeExist", "cUpdate_Click")
   
   frmImportControls!cCodeExist.Caption = "Update"
   Unload Me

   Call CallStackPop

End Sub



