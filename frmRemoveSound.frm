VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   Caption         =   "Scanner noise cancellation tool"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   FillColor       =   &H00808000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLogFile 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkNoMatch 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtTestRange 
      Height          =   285
      Left            =   5040
      TabIndex        =   18
      Text            =   "5"
      Top             =   1800
      Width           =   375
   End
   Begin MSComDlg.CommonDialog objFd 
      Left            =   1320
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtTopPerc 
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Text            =   "50"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtUpsampleFactor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5040
      TabIndex        =   13
      Text            =   "2"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtRamps 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Text            =   "10"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtThresholdForSound 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Text            =   "-10"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtChunkSize 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Text            =   "45"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtTRWithin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "0.01"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtTR 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "1.1"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Add to log file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   23
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Don't test for match, just use TR estimate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   21
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Test range for each pulse (samples)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   19
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Generate mean from most typical %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3120
      TabIndex        =   17
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "OPTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Upsampling factor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3840
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ramps (ms)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Threshold (dB relative to full scale)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Chunk size (s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "mean"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "within"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dx7 As New DirectX7
Dim ds As DirectSound
Dim dsb As DirectSoundBuffer
Dim bufferDesc As DSBUFFERDESC
Dim Format As WAVEFORMATEX
Dim intSnd() As Integer
Dim intSndSq() As Long
Const numDS = 4
Dim intDividers(1 To numDS) As Integer
Dim lngSamples(1 To numDS) As Long
Dim upsamplefactor As Integer






Function Search(buffer() As Integer, buffersq() As Long, intDSNum As Integer, start As Long, ln As Long, min As Long, max As Long, correl() As Single)
ReDim correl(min To max) As Single
Dim i, j, k As Long

Dim tot As Double
Dim tot1 As Double
Dim tot2 As Double
Dim lag As Long

For lag = min To max
    tot = 0
    tot1 = 0
    tot2 = 0
    For i = 0 To ln - 1
        tot = tot + CLng(buffer(start + i, intDSNum)) * buffer(start + lag + i, intDSNum)
        tot1 = tot1 + buffersq(start + i, intDSNum)
        tot2 = tot2 + buffersq(start + lag + i, intDSNum)
    Next
    tot = tot / ln
    tot1 = tot1 / ln
    tot2 = tot2 / ln
    
    If (tot1 = 0 Or tot2 = 0) Then
        correl(lag) = 0
    Else
        correl(lag) = tot / Sqr(CDbl(tot1) * tot2)
    End If
Next

End Function

Function SearchFixed(buffer2() As Long, buffersq2() As Long, buffer() As Integer, buffersq() As Long, intDSNum As Integer, start As Long, ln As Long, min As Long, max As Long, correl() As Single)
ReDim correl(min To max) As Single
Dim i, j, k As Long

Dim tot As Double
Dim tot1 As Double
Dim tot2 As Double
Dim lag As Long

Dim start2 As Long

For lag = min To max
    tot = 0
    tot1 = 0
    tot2 = 0
    For i = 0 To ln - 1
        tot = tot + CLng(buffer2(i + 1)) * buffer(start + lag + i, intDSNum)
        tot1 = tot1 + buffersq2(i + 1)
        tot2 = tot2 + buffersq(start + lag + i, intDSNum)
    Next
    tot = tot / ln
    tot1 = tot1 / ln
    tot2 = tot2 / ln
    
    correl(lag) = tot / Sqr(tot1 * tot2)
Next

End Function


Private Sub Command1_Click()
Dim i, j, k As Long

Dim lngSamplesRaw As Long

Dim ff As Integer, ffo As Integer
Dim tmp As wavriff, header As formatwav

Command1.Enabled = False
Command2.Enabled = False
Form1.Enabled = False


objFd.DialogTitle = "Choose sound file(s)"
objFd.DefaultExt = ".wav"
objFd.Filter = "WAV files (.wav) |*.wav|All files (*.*) |*.*"
objFd.flags = cdlOFNAllowMultiselect Or cdlOFNLongNames
objFd.ShowOpen
Debug.Print objFd.FileName

Dim intPos As Integer, intLastPos As Integer
Dim strFN(512) As String
Dim strPath As String
Dim intNumfiles As Integer


intPos = InStr(1, objFd.FileName, " ", vbBinaryCompare)
If (intPos > 1) Then
    strPath = Left(objFd.FileName, intPos - 1)
    intLastPos = intPos
    Do
        intPos = InStr(intLastPos + 1, objFd.FileName, " ")
        If (intPos <= 0) Then Exit Do
        intNumfiles = intNumfiles + 1
        strFN(intNumfiles) = Mid(objFd.FileName, intLastPos + 1, intPos - intLastPos - 1)
        intLastPos = intPos
    Loop
    intNumfiles = intNumfiles + 1
    strFN(intNumfiles) = Mid(objFd.FileName, intLastPos + 1)
 Else
    intNumfiles = intNumfiles + 1
    strFN(intNumfiles) = objFd.FileName
 End If
 
Dim nf As Integer

For nf = 1 To intNumfiles
    StartLog strFN(nf)
    frmStatus.lblFile = "Now working on " & strPath & strFN(nf)
    frmStatus.Visible = True
    frmStatus.Refresh
    frmStatus.ProgressBar1.Value = 0
    
    ff = LoadWAVheader(Format, strPath & strFN(nf), tmp, header)
    
    Dim strOutfn As String
    
    strOutfn = Left(strPath & strFN(nf), Len(strPath & strFN(nf)) - 4)
    
    ffo = SaveWAVheader(Format, strOutfn & "_processed.wav", tmp, header)
    
    If (Format.nBlockAlign <> 2) Then
        MsgBox "Only 16 bit wav files are supported: please convert first."
        End
    End If
        
    If (Format.nChannels <> 1) Then
        MsgBox "Only mono files can be loaded: please convert first."
        End
    End If
    
    ' Lets load up the first 30s ish
    
    Dim lngSamplesToLoad As Long
    
    lngSamplesToLoad = Format.lSamplesPerSec * txtChunkSize
    lngSamplesRaw = lngSamplesToLoad
    ReDim intSndRaw(1 To lngSamplesRaw) As Integer
        
    Dim ramplen As Long, ramplen2 As Long
    upsamplefactor = txtUpsampleFactor
    
    ramplen = upsamplefactor * Format.lSamplesPerSec * txtRamps / 1000
    ramplen2 = Int(ramplen / 2)
    
    Dim lngSamplesSoFar As Long
    lngSamplesSoFar = 0
    
    Do
        Get #ff, , intSndRaw
        lngSamplesSoFar = lngSamplesSoFar + lngSamplesToLoad
        ' first, lets find the start
        
        Dim sngRmsbinsize As Single
        Dim intRMsbinsize As Long
        Dim intNumRMsbins As Long
        Dim sngRMS() As Single
        Dim intSoundStart As Long
        
        
        sngRmsbinsize = 1   '1 ms
        intRMsbinsize = Int(sngRmsbinsize * Format.lSamplesPerSec / 1000)
        intNumRMsbins = Int(lngSamplesRaw / intRMsbinsize)
        ReDim sngRMS(1 To intNumRMsbins)
        
        Dim dblRMStot As Double
        
        For i = 1 To intNumRMsbins
            dblRMStot = 0
            For j = 0 To intRMsbinsize - 1
                dblRMStot = dblRMStot + intSndRaw(1 + j + (i - 1) * intRMsbinsize) ^ 2
            Next
            If (dblRMStot = 0) Then
                sngRMS(i) = -99999
            Else
                sngRMS(i) = 10 * Log(Sqr(dblRMStot / intRMsbinsize) / 23170.48)
            End If
        Next
    
        ' Find first bin over RMS threshold
        For i = 1 To intNumRMsbins
            If (sngRMS(i) > txtThresholdForSound) Then
                AddToLog "Found start of sound at " & i * sngRmsbinsize
                Exit For
            End If
        Next
        If (i = (intNumRMsbins + 1)) Then
            intSoundStart = lngSamplesRaw
        Else
            intSoundStart = i * intRMsbinsize
        End If
        Dim intASample As Integer
        ' Now write the blank bit to file
        
        For i = 1 To intSoundStart - 1
            intASample = intSndRaw(i)
            Put #ffo, , intASample
        Next
        
    Loop Until intSoundStart < lngSamplesRaw
    
    
    
    ' So, now we have a full buffer of nice noise
    
    Const sincwindow = 3
    Const PI = 3.141592
    Dim intStartBase As Long
    intStartBase = Int(lngSamplesRaw * upsamplefactor / 2)
    ReDim intSnd(lngSamplesRaw * upsamplefactor, 1 To numDS)
    ReDim intSndSq(lngSamplesRaw * upsamplefactor, 1 To numDS)
    
    ' Make sinc factors
    Dim sincfactors() As Single
    ReDim sincfactors(1 To upsamplefactor - 1, -(sincwindow - 1) To sincwindow)
    Dim tot As Single, totmult As Single
    Dim factor As Single
    Dim sngDesiredPos As Single
    Dim intLower As Long
    
    Dim dist As Single
    
    For i = 1 To upsamplefactor - 1 'Don't need zero, as this is straight copy
        tot = 0
        For j = -(sincwindow - 1) To sincwindow
            dist = Abs(j - CSng(i) / upsamplefactor)
            sincfactors(i, j) = (0.54 + (0.46 * Cos(PI * dist / sincwindow))) * Sin(PI * dist) / (PI * dist)
            tot = tot + sincfactors(i, j)
        Next
        For j = -(sincwindow - 1) To sincwindow
            sincfactors(i, j) = sincfactors(i, j) / tot
        Next
    Next
    
    Dim lngSoundStart_Upsampled As Long
    
    lngSoundStart_Upsampled = -1
    
    Dim qq As Single
    
    Dim booNotFirst As Boolean
    
    
    Do
    ' Shift the sound back over the blank bit and load up new data
        frmStatus.ProgressBar1.Value = 2 + (98# * lngSamplesSoFar) / header.datalen
        Debug.Print "Progress " & frmStatus.ProgressBar1.Value & "%"
        frmStatus.Refresh
    
        If (intSoundStart > 1) Then
            For i = 1 To (lngSamplesRaw - intSoundStart + 1)
                intSndRaw(i) = intSndRaw(i + intSoundStart - 1)
            Next
            For i = (lngSamplesRaw - intSoundStart + 2) To lngSamplesRaw
                If (EOF(ff)) Then Exit For
                Get #ff, , intASample
                intSndRaw(i) = intASample
            Next
            lngSamplesSoFar = lngSamplesSoFar + intSoundStart - 2
            lngSamplesRaw = i - 1
        End If
        lngSamples(1) = lngSamplesRaw * upsamplefactor
        
               
        AddToLog "Upsampling by factor of " & upsamplefactor
        DoEvents
    
        Dim ind As Long, startind As Long
    
        ' Now is there ready-upsampled data we can use?
        If (lngSoundStart_Upsampled = -1) Then
            lngSoundStart_Upsampled = lngSamples(1) + 1
            Dim lngChunkAlreadyGot As Long
            lngChunkAlreadyGot = lngSamples(1) - lngSoundStart_Upsampled + 1
            If (lngChunkAlreadyGot > 0) Then lngChunkAlreadyGot = lngChunkAlreadyGot - sincwindow - 1 ' Avoid those dodgy samples at the end, and do them properly this time
            For i = 1 To lngChunkAlreadyGot
                intSnd(i, 1) = intSnd(i + lngSoundStart_Upsampled, 1)
                intSndSq(i, 1) = intSndSq(i + lngSoundStart_Upsampled, 1)
            Next
            startind = lngChunkAlreadyGot + 1
        Else
            startind = 1
        End If
            
        ind = startind
        For i = 1 + Int((startind - 1) / upsamplefactor) To lngSamplesRaw
            DoEvents
            For j = (startind - 1) Mod upsamplefactor To (upsamplefactor - 1)
                If (j = 0) Then
                    intSnd(ind, 1) = intSndRaw(i)
                Else
                    totmult = 0
                    For k = -(sincwindow - 1) To sincwindow
                        If ((k + i) >= 1 And (k + i) <= lngSamplesRaw) Then totmult = totmult + intSndRaw(k + i) * sincfactors(j, k)
                    Next
                    intSnd(ind, 1) = totmult
                End If
            '    Debug.Print ind, intSnd(ind, 1)
                intSndSq(ind, 1) = CLng(intSnd(ind, 1)) * intSnd(ind, 1)
                ind = ind + 1
            Next
        Next
            
        If Not booNotFirst Then
        
            intDividers(1) = 1
            
            AddToLog "Now finding sample rate."
            
            ' Now downsample to 10ms bins
            Dim intDivider As Integer
            Dim GetSR As Double
            
            DoDownSample 2, Int(0.01 * Format.lSamplesPerSec * upsamplefactor)
            DoDownSample 3, Int(0.001 * Format.lSamplesPerSec * upsamplefactor)
            DoDownSample 4, Int(0.0001 * Format.lSamplesPerSec * upsamplefactor)
            
            GetSR = GetSampleRate(2, CSng(txtTR), CSng(txtTRWithin))
            AddToLog "First TR estimate:" & GetSR
            
            GetSR = GetSampleRate(3, GetSR, 0.01)
            AddToLog "Second TR estimate:" & GetSR
            GetSR = GetSampleRate(4, GetSR, 0.001)
            AddToLog "Third TR estimate:" & GetSR
            GetSR = GetSampleRate(1, GetSR, 0.0001)
            AddToLog "Fourth TR estimate:" & GetSR
        End If
        
        ' Start in middle
        ' Advance one TR
        Dim correl() As Single
        Dim dblBase As Double, dblNext As Double
        Dim intDir As Integer
        Dim timemax As Long
        Dim maxcorrel As Single
        Dim window As Long
        Dim TRmin_samp As Long, TRmax_samp As Long
        
        
        Dim ScanTimes(1 To 99999) As Long
        Dim ScanCorrel(1 To 99999) As Single
        Dim NumScans As Integer
        
        NumScans = 1
        ScanTimes(1) = intStartBase
        ScanCorrel(1) = 1
        
        Dim maxtimemax As Long
        
        
        window = Int(1 * GetSR * Format.lSamplesPerSec * upsamplefactor)
        For intDir = -1 To 1 Step 2
            dblBase = intStartBase
            Do
                DoEvents
                dblNext = dblBase + intDir * GetSR * Format.lSamplesPerSec * upsamplefactor
                If (dblNext < 1 Or dblNext > (lngSamples(1) - window)) Then Exit Do
                If chkNoMatch Then
                    NumScans = NumScans + 1
                    ScanTimes(NumScans) = dblNext
                    ScanCorrel(NumScans) = 1#
                    dblBase = dblNext
                Else
                    TRmin_samp = dblNext - intStartBase - txtTestRange
                    TRmax_samp = dblNext - intStartBase + txtTestRange
                    Search intSnd, intSndSq, 1, intStartBase, window, TRmin_samp, TRmax_samp, correl()
                    maxcorrel = -9999
                    For j = TRmin_samp To TRmax_samp
                        If (correl(j) > maxcorrel) Then
                            timemax = j
                            maxcorrel = correl(j)
                        End If
                    Next
                    NumScans = NumScans + 1
                    ScanTimes(NumScans) = intStartBase + timemax
                    ScanCorrel(NumScans) = maxcorrel
                    dblBase = intStartBase + timemax
                End If
                AddToLog "Found scan at " & Round(ScanTimes(NumScans) / (Format.lSamplesPerSec * upsamplefactor), 4) & "s with correlation " & maxcorrel
            Loop Until 0
        Next
        
        
        ' Now sort scan times
        
        Dim booSwap As Boolean, temp
        
        booSwap = True
        Do While booSwap
            booSwap = False
            For i = 1 To NumScans - 1
                If (ScanTimes(i + 1) < ScanTimes(i)) Then
                    booSwap = True
                    temp = ScanTimes(i)
                    ScanTimes(i) = ScanTimes(i + 1)
                    ScanTimes(i + 1) = temp
                End If
            Next
        Loop
        
        ' Find max interval
        maxtimemax = 0
        For i = 1 To NumScans - 1
            If ((ScanTimes(i + 1) - ScanTimes(i)) > maxtimemax) Then maxtimemax = ScanTimes(i + 1) - ScanTimes(i)
        Next
        
        
        ' Now sort scan times
        
        
        booSwap = True
        Do While booSwap
            booSwap = False
            For i = 1 To NumScans - 1
                If (ScanTimes(i + 1) < ScanTimes(i)) Then
                    booSwap = True
                    temp = ScanTimes(i)
                    ScanTimes(i) = ScanTimes(i + 1)
                    ScanTimes(i + 1) = temp
                End If
            Next
        Loop
        
        
        Dim neighb, neighb2
        neighb = 15
        neighb2 = Int(neighb / 2)
        
        maxtimemax = 0
        For i = 1 To NumScans - 1
            If ((ScanTimes(i + 1) - ScanTimes(i)) > maxtimemax) Then maxtimemax = ScanTimes(i + 1) - ScanTimes(i)
        Next
        
        
        ' Now calculate mean noise
        Dim n
        ReDim scansound(1 - ramplen To maxtimemax + ramplen)
            
        ' If very few then use mean from last time
        
        If (Not booNotFirst And NumScans < 3) Then
            MsgBox "Fewer than three scans were detected in the first chunk - choose a larger chunk size?"
            Exit Sub
        End If
        If (NumScans >= 3) Then
            For j = 1 - ramplen To maxtimemax + ramplen
                scansound(j) = 0
            Next
            n = 0
            ' Don't subtract from final one as will use this as a base for next window - and don't include in mean if second time round
            
            Dim offset As Integer
            If (booNotFirst) Then
                offset = 1
            Else
                offset = 0
            End If
            For i = offset + 1 To NumScans 'intMeanBase To intMeanBase + neighb - 1
                'If (i <> k) Then
                    DoEvents
                    n = n + 1
                    For j = 1 - ramplen To maxtimemax + ramplen
                        scansound(j) = scansound(j) + intSnd(j + ScanTimes(i) - 1, 1)
                    Next
                'End If
            Next
            For j = 1 - ramplen To maxtimemax + ramplen
                scansound(j) = scansound(j) / n
            Next
            
          ' Now, lets test which are the most typical samples (hopefully these will be the ones without speech)
          
          Dim sngRMSdiff() As Single
          Dim dblTot As Double
          ReDim sngRMSdiff(offset + 1 To NumScans)
          For i = offset + 1 To NumScans
                dblTot = 0
                For j = 1 To maxtimemax
                    dblTot = dblTot + (intSnd(j + ScanTimes(i) - 1, 1) - scansound(j)) ^ 2
                Next
                sngRMSdiff(i) = Sqr(dblTot / maxtimemax)
          Next
          
          Dim intLabels() As Integer
          ReDim intLabels(offset + 1 To NumScans)
          
          For i = offset + 1 To NumScans
            intLabels(i) = i
          Next
        Dim sngTemp As Single, intTemp As Integer
        
            booSwap = True
            Do While booSwap
              booSwap = False
              For i = offset + 1 To NumScans - 1
                  If (sngRMSdiff(i) > sngRMSdiff(i + 1)) Then
                      sngTemp = sngRMSdiff(i)
                      sngRMSdiff(i) = sngRMSdiff(i + 1)
                      sngRMSdiff(i + 1) = sngTemp
                      intTemp = intLabels(i)
                      intLabels(i) = intLabels(i + 1)
                      intLabels(i + 1) = intTemp
                      booSwap = True
                  End If
              Next
            Loop
              
            Dim intNuminMean As Integer
            
            intNuminMean = Int(txtTopPerc / 100# * (NumScans - (offset + 1)))
            
            AddToLog "Using most typical " & intNuminMean & " scans to generate mean for subtraction."
            
            If (intNuminMean >= 3) Then
            ' Now recalculate mean noise using only our favourites
             
             For j = 1 - ramplen To maxtimemax + ramplen
                 scansound(j) = 0
             Next
             n = 0
             
             For i = offset + 1 To offset + intNuminMean
                     DoEvents
                     n = n + 1
                     For j = 1 - ramplen To maxtimemax + ramplen
                         scansound(j) = scansound(j) + intSnd(j + ScanTimes(intLabels(i)) - 1, 1)
                     Next
             Next
             
             For j = 1 - ramplen To maxtimemax + ramplen
                 scansound(j) = scansound(j) / n
             Next
            End If
        End If
       ' For j = 1 To lngSamples(1)
      '      intSndCopy(j) = intSnd(j, 1)
      '  Next
        
        
           
        Dim ln As Long
        Dim rampfactor As Single
        
        
        
        For k = 1 To NumScans - 1
            DoEvents
            If (k = NumScans) Then
                ln = maxtimemax
            Else
                ln = ScanTimes(k + 1) - ScanTimes(k)
            End If
            For j = 1 To ramplen - 1
                rampfactor = (1 - Cos((j - 1) / ramplen * PI)) / 2
                intSnd(j + ScanTimes(k) - 1, 1) = intSnd(j + ScanTimes(k) - 1, 1) - rampfactor * scansound(j)
            Next
            
            For j = ramplen To ln
                intSnd(j + ScanTimes(k) - 1, 1) = intSnd(j + ScanTimes(k) - 1, 1) - scansound(j)
            Next
            
            For j = ln + 1 To ln + ramplen
                rampfactor = (1 + Cos((j - ln) / ramplen * PI)) / 2
                intSnd(j + ScanTimes(k) - 1, 1) = intSnd(j + ScanTimes(k) - 1, 1) - rampfactor * scansound(j)
            Next
        Next
        
        Dim lngtot As Long
        
        For j = 1 To lngSamplesRaw
            lngtot = 0
            For k = 0 To upsamplefactor - 1
                lngtot = lngtot + intSnd((j - 1) * upsamplefactor + 1, 1)
            Next
            intSndRaw(j) = lngtot / upsamplefactor
        Next
        
        intSoundStart = Int((ScanTimes(NumScans)) / upsamplefactor)
        lngSoundStart_Upsampled = intSoundStart * upsamplefactor
        
        If (EOF(ff)) Then
            For i = 1 To lngSamplesRaw
                intASample = intSndRaw(i)
                Put #ffo, , intASample
            Next
            Exit Do
        Else
    '        intASample = 32000
     '       Put #ffo, , intASample
            For i = 1 To intSoundStart - 1
                intASample = intSndRaw(i)
                Put #ffo, , intASample
            Next
            intStartBase = ScanTimes(NumScans) - lngSoundStart_Upsampled + 1
        End If
        
        booNotFirst = True
    Loop
    Close #ff
    Close #ffo
    
    AddToLog "File done."
    frmStatus.Visible = False
    Refresh
    Next
    AddToLog "All done."
Command1.Enabled = True
Command2.Enabled = True
Form1.Enabled = True

    End Sub
    
Sub StartLog(strFN As String)
If (chkLogFile) Then
    Dim ff As Integer
    ff = FreeFile
    Open "scannersound_log.txt" For Append As #ff
    Print #ff, "*** NEW SESSION: FILE " & strFN & " AT " & Now()
    Close #ff
End If
End Sub

Sub AddToLog(strMsg As String)
frmStatus.lblStatus = strMsg
frmStatus.Refresh
If (chkLogFile) Then
    Dim ff As Integer
    ff = FreeFile
    Open "scannersound_log.txt" For Append As #ff
    Print #ff, strMsg
    Close #ff
End If
End Sub

Function DoDownSample(intDSNum As Integer, intDivider As Integer)
Dim sngTot As Single
Dim i, j, k As Long
intDividers(intDSNum) = intDivider
lngSamples(intDSNum) = Int(lngSamples(1) / intDivider)
For i = 0 To lngSamples(intDSNum) - 1
    sngTot = 0#
    For j = 0 To intDivider - 1
        sngTot = sngTot + Abs(intSnd(i * intDivider + j + 1, 1))
    Next
    intSnd(i + 1, intDSNum) = Int(sngTot / intDivider)
    intSndSq(i + 1, intDSNum) = CLng(intSnd(i + 1, intDSNum)) * intSnd(i + 1, intDSNum)
Next
End Function


Function GetSampleRate(intDSNum As Integer, sngTR As Double, sngRange As Single)
Dim i, j, k As Long
Dim mean As Single, sd As Single
Dim intDivider As Integer
intDivider = intDividers(intDSNum)

' And calc correlations
Dim TRmin_samp As Long, TRmax_samp As Long

' First, work out exact TR

Dim lngStart As Long

Dim intNumVals As Long
Dim measurements(1 To 100) As Single

Dim window As Long
Dim correl() As Single

Dim maxcorrel As Single
Dim timemax As Long

Dim tot As Double
Dim totsq As Double

window = Int(sngTR * 3 * Format.lSamplesPerSec * upsamplefactor / intDivider)

Dim intMinVals As Long

If (intDSNum = 1 And chkNoMatch) Then
    intMinVals = 20
Else
    intMinVals = 2
End If

For i = 1 To 100
    DoEvents
    TRmin_samp = CLng((sngTR - sngRange) * Format.lSamplesPerSec * upsamplefactor / intDivider)
    TRmax_samp = CLng((sngTR + sngRange) * Format.lSamplesPerSec * upsamplefactor / intDivider)
    lngStart = Int(Rnd() * (lngSamples(intDSNum) - window - TRmax_samp - 1)) + 1
    
    Search intSnd, intSndSq, intDSNum, lngStart, window, TRmin_samp, TRmax_samp, correl
    
    maxcorrel = -9999
    For j = TRmin_samp To TRmax_samp
        If (correl(j) > maxcorrel) Then
            timemax = j
            maxcorrel = correl(j)
        End If
    Next
    Debug.Print "Maximum correlation:", maxcorrel
    intNumVals = intNumVals + 1
    measurements(intNumVals) = timemax

    
    If (intNumVals > intMinVals) Then
        tot = 0
        totsq = 0
        For j = 1 To intNumVals
            tot = tot + measurements(j)
            totsq = totsq + measurements(j) * measurements(j)
        Next
        mean = tot / intNumVals
        sd = Sqr((intNumVals * totsq - (tot * tot)) / (intNumVals * (intNumVals - 1)))
        If (sd < 1) Then Exit For
    End If
Next
GetSampleRate = mean / Format.lSamplesPerSec * intDivider / upsamplefactor


End Function


Private Sub Form_Terminate()
End
End Sub


