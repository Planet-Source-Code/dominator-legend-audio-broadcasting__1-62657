VERSION 5.00
Begin VB.UserControl FullDoplex 
   AccessKeys      =   "AudioStreamer"
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   900
   ScaleWidth      =   3945
   ToolboxBitmap   =   "AudioStreamer.ctx":0000
   Begin VB.Timer VirtualEvent 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   30
      Top             =   30
   End
   Begin VB.Label Lbl_Interface 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Audio FullDoplex Module"
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
      Left            =   1350
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image Img_Interface 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   0
      Picture         =   "AudioStreamer.ctx":0312
      Top             =   0
      Width           =   3885
   End
End
Attribute VB_Name = "FullDoplex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private DX_Main                         As New DirectX7
Private DX_Reader                       As DirectSoundCapture
Private DX_ReaderBuffer                 As DirectSoundCaptureBuffer
Private DX_Writer                       As DirectSound
Private DX_WriterBuffer                 As DirectSoundBuffer
Private WAVBuffer()                     As Byte
Private SamplesPerSec                   As Long
Private Channels                        As Byte
Private BitsPerSample                   As Integer
Private Sub UserControl_Resize()
Attribute UserControl_Resize.VB_Description = "Events occure when the control resized."
    UserControl.Width = Img_Interface.Width
    UserControl.Height = Img_Interface.Height
End Sub
Public Sub InitStreamer()
Attribute InitStreamer.VB_Description = "Method which init the stream object and start recording and playing"
    If DX_ReaderBuffer Is Nothing Or DX_WriterBuffer Is Nothing Then Call CreateBuffer
    DX_ReaderBuffer.start DSBPLAY_LOOPING
    DX_WriterBuffer.Play DSBPLAY_LOOPING
    VirtualEvent.Enabled = True
End Sub
Private Function CreateBuffer() As String
Attribute CreateBuffer.VB_Description = "Method which create the a buffer, which used to hold the captured and played data"
On Error GoTo ErrorHandel:
    Dim RBufferDescription           As DSCBUFFERDESC
    Dim RWAVFormat                   As WAVEFORMATEX
    Dim PBufferDescription           As DSBUFFERDESC
    Dim PWAVFormat                   As WAVEFORMATEX
    Set DX_Reader = DX_Main.DirectSoundCaptureCreate("")
    Set DX_Writer = DX_Main.DirectSoundCreate("")
    DX_Writer.SetCooperativeLevel UserControl.hWnd, DSSCL_PRIORITY
    RWAVFormat = SetWavFormat(SamplesPerSec, Channels, BitsPerSample)
    PWAVFormat = SetWavFormat(SamplesPerSec, Channels, BitsPerSample)
    RBufferDescription.fxFormat = RWAVFormat
    RBufferDescription.lBufferBytes = RWAVFormat.lAvgBytesPerSec
    RBufferDescription.lFlags = DSCBCAPS_WAVEMAPPED
    PBufferDescription.lBufferBytes = PWAVFormat.lAvgBytesPerSec
    PBufferDescription.lFlags = DSBCAPS_CTRLPOSITIONNOTIFY Or DSBCAPS_GLOBALFOCUS Or DSBCAPS_GETCURRENTPOSITION2
    Set DX_ReaderBuffer = DX_Reader.CreateCaptureBuffer(RBufferDescription)
    Set DX_WriterBuffer = DX_Writer.CreateSoundBuffer(PBufferDescription, PWAVFormat)
Exit Function
ErrorHandel:
CreateBuffer = "Error Source: " & Err.Source & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description
End Function
Private Function SetWavFormat(ByVal WSamplesPerSec As Long, ByVal WChannels As Byte, ByVal WBitsPerSample As Integer) As WAVEFORMATEX
Attribute SetWavFormat.VB_Description = "Method which used to determine the captured quality."
    SetWavFormat.nFormatTag = 1
    SetWavFormat.nChannels = WChannels
    SetWavFormat.lSamplesPerSec = WSamplesPerSec
    SetWavFormat.nBitsPerSample = WBitsPerSample
    SetWavFormat.nBlockAlign = WChannels * WBitsPerSample / 8
    SetWavFormat.lAvgBytesPerSec = SetWavFormat.lSamplesPerSec * SetWavFormat.nBlockAlign
    SetWavFormat.nSize = 0
End Function
Public Sub SetWavQuality(ByVal WSamplesPerSec As Long, ByVal WChannels As Byte, ByVal WBitsPerSample As Integer)
Attribute SetWavQuality.VB_Description = "Method which used to determine the captured quality."
    SamplesPerSec = WSamplesPerSec
    Channels = WChannels
    BitsPerSample = WBitsPerSample
End Sub
Public Sub DestroyObjects()
Attribute DestroyObjects.VB_Description = "Method which clear the memory, and destroy created objects."
Attribute DestroyObjects.VB_UserMemId = 0
    Set DX_Main = Nothing
    Set DX_Reader = Nothing
    Set DX_ReaderBuffer = Nothing
    Set DX_Writer = Nothing
    Set DX_WriterBuffer = Nothing
End Sub
Public Sub StopStreaming()
Attribute StopStreaming.VB_Description = "Method which stop the stream, but don't destroy the objects."
    DX_ReaderBuffer.Stop
    DX_WriterBuffer.Stop
End Sub
Private Sub VirtualEvent_Timer()
Attribute VirtualEvent_Timer.VB_Description = "An events occure periodically to retive to us the captured data, in this event the data are copied from the host buffer to the client buffer through the copybuffer method."
   Call CopyBuffers
End Sub
Private Sub CopyBuffers()
Attribute CopyBuffers.VB_Description = "Method which copy the recorded data from the local buffer to the remote buffer."
    On Error Resume Next
    Dim Cursors                     As DSCURSORS
    DX_ReaderBuffer.GetCurrentPosition Cursors
    ReDim WAVBuffer(Cursors.lWrite + 1)
    DX_ReaderBuffer.ReadBuffer 0, UBound(WAVBuffer), WAVBuffer(0), DSCBLOCK_DEFAULT
    DX_WriterBuffer.WriteBuffer 0, UBound(WAVBuffer), WAVBuffer(0), DSBLOCK_DEFAULT
End Sub
