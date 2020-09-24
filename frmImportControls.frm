VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmImportControls 
   Caption         =   "Importing form (hidden)"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7512
   LinkTopic       =   "Form1"
   ScaleHeight     =   6672
   ScaleWidth      =   7512
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox cMailDate 
      Height          =   372
      Left            =   1320
      TabIndex        =   8
      Top             =   5160
      Width           =   972
   End
   Begin VB.TextBox cSubmitted 
      Height          =   372
      Left            =   1320
      TabIndex        =   7
      Top             =   4680
      Width           =   972
   End
   Begin VB.TextBox cCompatibility 
      Height          =   372
      Left            =   1320
      TabIndex        =   6
      Top             =   4200
      Width           =   972
   End
   Begin VB.TextBox cURL 
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   972
   End
   Begin VB.TextBox cDes 
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   972
   End
   Begin VB.TextBox cLevel 
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   972
   End
   Begin VB.TextBox cCategory 
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   972
   End
   Begin VB.TextBox cTitle 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   972
   End
   Begin MSFlexGridLib.MSFlexGrid cMessageText 
      Height          =   3972
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7212
      _ExtentX        =   12721
      _ExtentY        =   7006
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label cCodeExist 
      BorderStyle     =   1  'Fixed Single
      Height          =   372
      Left            =   2880
      TabIndex        =   9
      Top             =   4200
      Width           =   1212
   End
End
Attribute VB_Name = "frmImportControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

