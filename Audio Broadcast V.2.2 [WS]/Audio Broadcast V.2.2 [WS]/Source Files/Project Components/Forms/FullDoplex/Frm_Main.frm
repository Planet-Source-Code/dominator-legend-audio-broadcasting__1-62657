VERSION 5.00
Begin VB.Form Frm_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio FullDoplex"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   FillColor       =   &H00404040&
   ForeColor       =   &H00808080&
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin AudioFullDoplex.FullDoplex Doplexer 
      Left            =   1170
      Top             =   150
      _ExtentX        =   6853
      _ExtentY        =   1376
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dominator Legend Audio FullDoplex"
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
      Width           =   5835
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 2005 Dominator Legend, Audio FullDoplex Application"
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
      Width           =   4020
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dominator Legend Audio FullDoplex"
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
      Width           =   5835
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Attribute Form_Load.VB_Description = "Main method to load the form"
    Call Doplexer.SetWavQuality(48000, 2, 16)
    Call Doplexer.InitStreamer
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Doplexer.StopStreaming
    Call Doplexer.DestroyObjects
End Sub
