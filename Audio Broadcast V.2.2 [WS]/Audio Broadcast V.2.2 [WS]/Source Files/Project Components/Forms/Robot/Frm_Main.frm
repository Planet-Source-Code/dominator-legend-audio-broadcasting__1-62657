VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Frm_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Robot Broadcaster"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   FillColor       =   &H00404040&
   ForeColor       =   &H00808080&
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Frm_Main.frx":29C12
   ScaleHeight     =   1065
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin Robot.AudioStreamer Streamer 
      Left            =   0
      Top             =   0
      _ExtentX        =   6853
      _ExtentY        =   1376
   End
   Begin ComctlLib.ImageList ToHandelTheManifestError 
      Left            =   -60
      Top             =   1110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dominator Legend Robot Streamer"
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
      Width           =   5730
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 2005 Dominator Legend, Streamer Application"
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
      Width           =   3495
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dominator Legend Robot Streamer"
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
      Width           =   5730
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Call Streamer.SetNetworkParameters(1, 19840, 19841, "127.0.0.1")
    Rem -> If we increase the quality we got an error cause winsocket have a small buffer
    Call Streamer.SetWavQuality(8000, 1, 8)
    Call Streamer.InitStreamer
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call Streamer.StopStreaming
    Call Streamer.DestroyObjects
End Sub
