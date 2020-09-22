VERSION 5.00
Begin VB.Form frmTEST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DLL Test Form"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPULL 
      Caption         =   "&Pull my finger."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblLINK 
      Caption         =   "http://createafart.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label lblLABEL 
      Caption         =   ".WAV's supplied by:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label lblTITLE 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fart Generator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image imgFINGER 
      Height          =   615
      Left            =   120
      Picture         =   "frmTEST.frx":0000
      Top             =   0
      Width           =   945
   End
   Begin VB.Shape shpTOP 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   -600
      Top             =   -120
      Width           =   3975
   End
End
Attribute VB_Name = "frmTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdPULL_Click()
    On Error Resume Next
    Dim obj As FARTRES.FART
    Set obj = New FARTRES.FART
    obj.RandomFart
End Sub

Private Sub lblLINK_Click()
    On Error Resume Next
    ExecuteLink "http://www.createafart.com/index.asp"
End Sub
