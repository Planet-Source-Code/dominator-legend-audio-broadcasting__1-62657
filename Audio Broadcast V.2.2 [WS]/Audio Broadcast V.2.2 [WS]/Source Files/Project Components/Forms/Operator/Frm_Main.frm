VERSION 5.00
Begin VB.Form Frm_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operator"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   FillColor       =   &H00404040&
   ForeColor       =   &H00808080&
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin Operator.AudioReceiver Receiver 
      Left            =   0
      Top             =   0
      _ExtentX        =   6853
      _ExtentY        =   1376
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dominator Legend Operator Simulator"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D09764&
      Height          =   510
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   240
      Width           =   6300
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 2005 Dominator Legend, Operator Simulator Application"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D09764&
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   0
      Top             =   720
      Width           =   4200
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dominator Legend Operator Simulator"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006B2401&
      Height          =   510
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   255
      Width           =   6300
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Call Receiver.SetNetworkParameters(1, 19841)
    Call Receiver.SetWavQuality(8000, 1, 8)
    Call Receiver.InitReceiver
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call Receiver.StopStreaming
    Call Receiver.DestroyObjects
End Sub
