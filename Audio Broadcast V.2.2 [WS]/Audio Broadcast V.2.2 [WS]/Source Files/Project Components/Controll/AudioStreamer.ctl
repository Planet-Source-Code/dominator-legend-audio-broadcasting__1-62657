VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl AudioStreamer 
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
      Left            =   450
      Top             =   30
   End
   Begin MSWinsockLib.Winsock Streamer 
      Left            =   30
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Lbl_Interface 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Audio Transmitter Module"
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
Attribute VB_Name = "AudioStreamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private DX_Main                         As New DirectX7
Private DX_Reader                       As DirectSoundCapture
Private DX_ReaderBuffer                 As DirectSoundCaptureBuffer
Private WAVBuffer()                     As Byte
Private SamplesPerSec                   As Long
Private Channels                        As Byte
Private BitsPerSample                   As Integer
Private Const PacketSize                As Long = 9000
Private Sub UserControl_Resize()
    UserControl.Width = Img_Interface.Width
    UserControl.Height = Img_Interface.Height
End Sub
Public Sub InitStreamer()
    If DX_ReaderBuffer Is Nothing Then Call CreateBuffer
    DX_ReaderBuffer.Start DSBPLAY_LOOPING
    VirtualEvent.Enabled = True
End Sub
Private Function CreateBuffer() As String
On Error GoTo ErrorHandel:
    Dim BufferDescription           As DSCBUFFERDESC
    Dim WAVFormat                   As WAVEFORMATEX
    Set DX_Reader = DX_Main.DirectSoundCaptureCreate("")
    WAVFormat = SetWavFormat(SamplesPerSec, Channels, BitsPerSample)
    BufferDescription.fxFormat = WAVFormat
    BufferDescription.lBufferBytes = WAVFormat.lAvgBytesPerSec
    BufferDescription.lFlags = DSCBCAPS_WAVEMAPPED
    Set DX_ReaderBuffer = DX_Reader.CreateCaptureBuffer(BufferDescription)
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
    Set DX_Reader = Nothing
    Set DX_ReaderBuffer = Nothing
End Sub
Public Sub StopStreaming()
    DX_ReaderBuffer.Stop
End Sub
Private Sub VirtualEvent_Timer()
   Call CopyBuffers
End Sub
Private Sub CopyBuffers()
    On Error Resume Next
    Dim Cursors                     As DSCURSORS
    DX_ReaderBuffer.GetCurrentPosition Cursors
    ReDim WAVBuffer(Cursors.lWrite + 1)
    DX_ReaderBuffer.ReadBuffer 0, UBound(WAVBuffer), WAVBuffer(0), DSCBLOCK_DEFAULT
    Streamer.SendData WAVBuffer
    'Cause winsocket need smaller packets i devide large packets into smaller one and send it
    'but this method dosn't work, the sound stay cutting. :(
    'Call SendPackets(PacketSize, WAVBuffer)
End Sub
Public Sub SetNetworkParameters(ByVal Protocol As Byte, ByVal LocalPort As Long, ByVal RemotePort As Long, ByVal RemoteIP As String)
    Streamer.Close
    Streamer.Protocol = Protocol
    Streamer.LocalPort = LocalPort
    Streamer.RemotePort = RemotePort
    Streamer.RemoteHost = RemoteIP
End Sub
Public Function CopyArray(ByRef Source() As Byte, ByRef Destenation() As Byte, ByVal Start As Long, ByVal Length As Long)
    Dim ExtElement      As Long
    Dim IntElement      As Long: IntElement = 0
    ReDim Destenation(Length)
    For ExtElement = Start To Length
        Destenation(IntElement) = Source(ExtElement)
        IntElement = IntElement + 1
    Next ExtElement
End Function
Public Function SendPackets(ByRef PacketSize As Long, ByRef Source() As Byte)
    Dim SourceLength        As Long
    Dim Curser              As Long
    Dim Element             As Long
    Dim Packet()            As Byte
    SourceLength = UBound(Source)
    Curser = 0: Element = 0
    Do While SourceLength > PacketSize
        Call CopyArray(Source, Packet, Curser, PacketSize)
        Streamer.SendData Source
        Curser = Curser + PacketSize
        SourceLength = SourceLength - PacketSize
        DoEvents
    Loop
    If SourceLength > 0 Then Call CopyArray(Source, Packet, Curser, SourceLength)
End Function
