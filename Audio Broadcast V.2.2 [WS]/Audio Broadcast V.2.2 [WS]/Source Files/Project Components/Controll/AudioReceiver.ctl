VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl AudioReceiver 
   AccessKeys      =   "AudioReceiver"
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   870
   ScaleWidth      =   3960
   ToolboxBitmap   =   "AudioReceiver.ctx":0000
   Begin VB.Timer Checker 
      Interval        =   2000
      Left            =   420
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Receiver 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Lbl_Interface 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Audio Receiver Module"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   2205
   End
   Begin VB.Image Img_Interface 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   0
      Picture         =   "AudioReceiver.ctx":0312
      Top             =   0
      Width           =   3885
   End
End
Attribute VB_Name = "AudioReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private DX_Main                         As New DirectX7
Private DX_Writer                       As DirectSound
Private DX_WriterBuffer                 As DirectSoundBuffer
Private WAVBuffer()                     As Byte
Private SamplesPerSec                   As Long
Private Channels                        As Byte
Private BitsPerSample                   As Integer
Private DataLength                      As Long
Private DataLengthTemp                  As Long
Private Match                           As Boolean
Private Sub UserControl_Resize()
    UserControl.Width = Img_Interface.Width
    UserControl.Height = Img_Interface.Height
End Sub
Public Sub InitReceiver()
    If DX_WriterBuffer Is Nothing Then Call CreateBuffer
    DX_WriterBuffer.Play DSBPLAY_LOOPING
End Sub
Private Function CreateBuffer() As String
On Error GoTo ErrorHandel:
    Dim BufferDescription           As DSBUFFERDESC
    Dim WAVFormat                   As WAVEFORMATEX
    Set DX_Writer = DX_Main.DirectSoundCreate("")
    DX_Writer.SetCooperativeLevel UserControl.hWnd, DSSCL_PRIORITY
    WAVFormat = SetWavFormat(SamplesPerSec, Channels, BitsPerSample)
    BufferDescription.lBufferBytes = WAVFormat.lAvgBytesPerSec
    BufferDescription.lFlags = DSBCAPS_CTRLPOSITIONNOTIFY Or DSBCAPS_GLOBALFOCUS Or DSBCAPS_GETCURRENTPOSITION2
    Set DX_WriterBuffer = DX_Writer.CreateSoundBuffer(BufferDescription, WAVFormat)
Exit Function
ErrorHandel:
CreateBuffer = "Error Source: " & Err.Source & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description
End Function
Private Function SetWavFormat(ByVal WSamplesPerSec As Long, ByVal WChannels As Byte, ByVal WBitsPerSample As Integer) As WAVEFORMATEX
    SetWavFormat.nFormatTag = WAVE_FORMAT_PCM
    SetWavFormat.nChannels = WChannels
    SetWavFormat.lSamplesPerSec = WSamplesPerSec
    SetWavFormat.nBitsPerSample = WBitsPerSample
    SetWavFormat.nBlockAlign = WChannels * WBitsPerSample / 8
    SetWavFormat.lAvgBytesPerSec = SetWavFormat.lSamplesPerSec * SetWavFormat.nBlockAlign
    SetWavFormat.nSize = 0
End Function
Public Sub SetWavQuality(ByVal WSamplesPerSec As Long, ByVal WChannels As Byte, ByVal WBitsPerSample As Integer)
    SamplesPerSec = WSamplesPerSec
    Channels = WChannels
    BitsPerSample = WBitsPerSample
End Sub
Public Sub DestroyObjects()
    Set DX_Main = Nothing
    Set DX_Writer = Nothing
    Set DX_WriterBuffer = Nothing
End Sub
Public Sub StopStreaming()
    DX_WriterBuffer.Stop
End Sub
Private Sub Receiver_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Receiver.GetData WAVBuffer, vbArray + vbByte, bytesTotal
    DX_WriterBuffer.WriteBuffer 0, UBound(WAVBuffer), WAVBuffer(0), DSBLOCK_DEFAULT
    'DataLength = bytesTotal
End Sub
'Private Sub Checker_Timer()
'    If DataLengthTemp = DataLength Then
'        StopStreaming
'    Else
'        DX_WriterBuffer.Play DSBPLAY_LOOPING
'        DataLengthTemp = DataLength
'    End If
'End Sub
Private Sub Receiver_Close()
    StopStreaming
End Sub
Public Sub SetNetworkParameters(ByVal Protocol As Byte, ByVal LocalPort As Long)
    Receiver.Close
    Receiver.Protocol = Protocol
    Receiver.LocalPort = LocalPort
    Receiver.Bind LocalPort
End Sub
