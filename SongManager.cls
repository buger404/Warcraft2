VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SongManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim SongHandle As Long
Dim FFTData(512) As Single
Dim FFTPeacks(128) As Long
Dim FFTFallOff(128) As Long
Public Rate As Long
Enum PlayStates
        Stopped = 0
        Playing = 1
        Stalled = 2
        Paused = 3
End Enum
Sub SetPlayRate(ByVal NRate As Single)
Rate = 44100 * NRate
BASS_ChannelSetAttribute SongHandle, BASS_ATTRIB_FREQ, 44100 * NRate
End Sub
Sub LoadMusic(ByVal Path As String)
  Call BASS_ChannelStop(SongHandle)
  SongHandle = BASS_StreamCreateFile(BASSFALSE, StrPtr(Path), 0, 0, 0)
  If SongHandle = 0 Then SongHandle = BASS_MusicLoad(BASSFALSE, StrPtr(Path), 0, 0, BASS_MUSIC_RAMPS Or BASS_MUSIC_POSRESET Or BASS_MUSIC_PRESCAN, 1)
  If SongHandle = 0 Then Msgbox "Music play failed .", , "Error"
  'BASS_ChannelSetAttribute SongHandle, BASS_ATTRIB_MUSIC_SPEED, 100
End Sub
Sub Play()
  Call BASS_ChannelPlay(SongHandle, BASSFALSE)
End Sub
Sub StopMusic()
  Call BASS_ChannelStop(SongHandle)
End Sub
Sub Pause()
 Call BASS_ChannelPause(SongHandle)
End Sub
Function PlayState() As PlayStates
PlayState = BASS_ChannelIsActive(SongHandle)
End Function
Function GetMusicBar() As Single()
On Error Resume Next
Dim i As Long, di As Long, j As Long, returns(512) As Single
Dim tmp As Long

 If BASS_ChannelIsActive(SongHandle) <> BASS_ACTIVE_PLAYING Then Exit Function
 
 Call BASS_ChannelGetData(SongHandle, FFTData(0), BASS_DATA_FFT1024)
 
 For i = 0 To UBound(FFTData) - 1
  di = Int(Abs(FFTData(i)) * 7500)
  j = i Mod 129
  If di > 1500 Then di = 1500
  If di >= FFTPeacks(j) Then
   FFTPeacks(j) = di
   Else
   FFTPeacks(j) = FFTPeacks(j) - 10
  End If
  If di >= FFTFallOff(j) Then
   FFTFallOff(j) = di
   Else
   FFTFallOff(j) = FFTFallOff(j) - 30
  End If
  If (1500 - FFTPeacks(j)) > 1500 Then FFTPeacks(j) = 0
  If (1500 - FFTFallOff(j)) > 1500 Then FFTFallOff(j) = 0
  returns(i) = FFTFallOff(j)
 Next
 GetMusicBar = returns
End Function

Private Sub Class_Initialize()
Rate = 44100
End Sub
