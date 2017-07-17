Attribute VB_Name = "Module1"
Dim booStop As Boolean


Type wavriff
char(1 To 4) As Byte
totallen As Long
wve(1 To 4) As Byte
End Type

Type formatwav
char(1 To 4) As Byte
formatlen As Long
dummy As Integer
channels As Integer
sr As Long
bytespersecond As Long
bytespersample As Integer
bitspersample As Integer
char2(1 To 4) As Byte
datalen As Long
End Type


Function LoadWAVheader(Format As WAVEFORMATEX, fn As String, tmp As wavriff, header As formatwav)

ff = FreeFile
Open fn For Binary As #ff
Get #ff, , tmp
Get #ff, , header

Format.lAvgBytesPerSec = header.bytespersecond
Format.lSamplesPerSec = header.sr
Format.nBitsPerSample = header.bitspersample
Format.nBlockAlign = header.bytespersample
Format.nChannels = header.channels

LoadWAVheader = ff
End Function

Function SaveWAVheader(Format As WAVEFORMATEX, fn As String, tmp As wavriff, header As formatwav)

header.bytespersecond = Format.lAvgBytesPerSec
header.sr = Format.lSamplesPerSec
header.bitspersample = Format.nBitsPerSample
header.bytespersample = Format.nBlockAlign
header.channels = Format.nChannels


ff = FreeFile
Open fn For Binary As #ff
Put #ff, , tmp
Put #ff, , header


SaveWAVheader = ff
End Function
