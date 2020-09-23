VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "Evan's CD Player"
   ClientHeight    =   3300
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8445
   ForeColor       =   &H00000000&
   Icon            =   "Cd Player1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Credits"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   2640
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Left            =   8280
      Top             =   3120
   End
   Begin VB.ComboBox TrackSelection 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   480
      Left            =   4680
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.VScrollBar Volume 
      Height          =   1455
      Left            =   7920
      MousePointer    =   4  'Icon
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6960
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   615
      Left            =   6000
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox TimeWindow 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   735
      Left            =   360
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Time"
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Play 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      ToolTipText     =   "Play"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Pause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Pause"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton stpButton 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      ToolTipText     =   "Stop"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton PreviousTrack 
      Caption         =   "Previous Track"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      ToolTipText     =   "Back One Song"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton NextTrack 
      Caption         =   "Next Track"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      ToolTipText     =   "Forward One Song"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Rewind 
      Caption         =   "Rewind"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Rewind"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton FastForward 
      Caption         =   "Fast Forward"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Fast Forward"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Eject 
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      ToolTipText     =   "Eject CD"
      Top             =   960
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   3720
   End
   Begin VB.Label lblTrackSelection 
      BackColor       =   &H00400000&
      Caption         =   "Track Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblVolume 
      BackColor       =   &H00400000&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   615
      Left            =   6120
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label TotalTrack 
      BackColor       =   &H80000006&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label TrackTime 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   1800
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************************
'***Program: CD Player  *****************************************************************************
'***Author: Evan Silich  ****************************************************************************
'***Created:  3/27/00   *****************************************************************************
'****************************************************************************************************
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Dim FastForwardSpeed As Long        ' seconds to seek for ff/rew
Dim Playing As Boolean                ' true if CD is currently playing
Dim CDLoad As Boolean                  ' true if CD is the the player
Dim TotalTracks As Integer              ' total tracks tracks on audio CD
Dim TrackLength() As String              ' array containing length of each track
Dim Track As Integer                     ' current track
Dim Minute As Integer                   ' current minute on track
Dim Second As Integer                  ' current second on track
Dim Command As String                 ' string to hold mci command strings
Dim hmixer As Long                   ' mixer handle
Dim volCtrl As MIXERCONTROL         ' Waveout volume control.

              
'Option Explicit

' Send a MCI command string
' If fShowError is true, display a message box on error
Private Function SendMCIString(Cmd As String, fShowError As Boolean) As Boolean
Static rc As Long               'return code
Static errStr As String * 400

rc = mciSendString(Cmd, 0, 0, hwnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function

Private Sub Close_Click()
SendMCIString "set cd door closed", True
Update
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Exit_Click()

SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
End
End Sub

Private Sub Form_Load()
Dim rc  As Long
Dim OK As Boolean
' Open the mixer with deviceID 0.
'
rc = mixerOpen(hmixer, 0, 0, 0, 0)
If MMSYSERR_NOERROR <> rc Then
    MsgBox "Could not open the mixer...You can still play your CD, but you volume control is disabled!", vbInformation, "Volume Control"
End If
'
' Get the waveout volume control.
'
OK = fGetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)
'
' If the function successfully gets the volume control,
' the maximum and minimum values are specified by
' lMaximum and lMinimum. Use them to set the scrollbar.
'
If OK Then
    With Volume
        .Max = volCtrl.lMinimum
        .Min = volCtrl.lMaximum \ 2
        .SmallChange = 1000
        .LargeChange = 1000
    End With
End If
    Left = (Screen.Width - Width) \ 2       'will center form
    Top = (Screen.Height - Height) \ 2      'will center form

Timer1.Interval = 500 'Change value depending On the speed of flahing.

If (App.PrevInstance = True) Then
    End
End If


Timer1.Enabled = False
FastForwardSpeed = 10
CDLoad = False


If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
    End
End If

SendMCIString "set cd time format tmsf wait", True
Timer1.Enabled = True
MsgBox ("Open CD rom Drive.")
SendMCIString "set cd door open", True  'sets cd door open
MsgBox ("Put your compact disk in the CD Rom drive and click Close.")

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close all MCI devices opened by this program
SendMCIString "close all", False
End Sub

' Play the CD
Private Sub Play_Click()
SendMCIString "play cd", True
Playing = True
End Sub

' Pause the CD
Private Sub Pause_Click()
SendMCIString "pause cd", True
Playing = False
Update
End Sub
' Eject the CD
Private Sub Eject_Click()
SendMCIString "set cd door open", True
Update
End Sub
' Fast forward
Private Sub FastForward_Click()
Dim e As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", e, Len(e), 0
If (Playing) Then
    Command = "play cd from " & CStr(CLng(e) + FastForwardSpeed * 1000)
Else
    Command = "seek cd to " & CStr(CLng(e) + FastForwardSpeed * 1000)
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub
' Rewind the CD
Private Sub Rewind_Click()
Dim e As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", e, Len(e), 0
If (Playing) Then
    Command = "play cd from " & CStr(CLng(e) - FastForwardSpeed * 1000)
Else
    Command = "seek cd to " & CStr(CLng(e) - FastForwardSpeed * 1000)
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub
' Forward track
Private Sub NextTrack_Click()
If (Track < TotalTracks) Then
    If (Playing) Then
        Command = "play cd from " & Track + 1
        SendMCIString Command, True
    Else
        Command = "seek cd to " & Track + 1
        SendMCIString Command, True
    End If
Else
    SendMCIString "seek cd to 1", True
End If
Update
End Sub
' Go to previous track
Private Sub PreviousTrack_Click()
Dim from As String
If (Minute = 0 And Second = 0) Then
    If (Track > 1) Then
        from = CStr(Track - 1)
    Else
        from = CStr(TotalTracks)
    End If
Else
    from = CStr(Track)
End If
If (Playing) Then
    Command = "play cd from " & from
    SendMCIString Command, True
Else
    Command = "seek cd to " & from
    SendMCIString Command, True
End If
Update
End Sub

' Update the display and state variables
Private Sub Update()
Static e As String * 30

' Check if CD is in the player
mciSendString "status cd media present", e, Len(e), 0
If (CBool(e)) Then
    ' Enable all the controls, get CD information
    If (CDLoad = False) Then
        mciSendString "status cd number of tracks wait", e, Len(e), 0
        TotalTracks = CInt(Mid$(e, 1, 2))
        Eject.Enabled = True
        
        ' If CD only has 1 track, then it's probably a data CD
        If (TotalTracks = 1) Then
            Exit Sub
        End If
        
        mciSendString "status cd length wait", e, Len(e), 0
        TotalTrack.Caption = "Tracks: " & TotalTracks & "  Total time: " & e
        ReDim TrackLength(1 To TotalTracks)
        Dim i As Integer
        For i = 1 To TotalTracks
            Command = "status cd length track " & i
            mciSendString Command, e, Len(e), 0
            TrackLength(i) = e
        Next
        Dim ts As Integer
        TrackSelection.Clear
        For ts = 1 To TotalTracks
        TrackSelection.AddItem ts
        Next ts
        TrackSelection.Text = TrackSelection.List(0)
        
        Play.Enabled = True
        Pause.Enabled = True
        FastForward.Enabled = True
        Rewind.Enabled = True
        NextTrack.Enabled = True
        PreviousTrack.Enabled = True
        stpButton.Enabled = True
        CDLoad = True
        SendMCIString "seek cd to 1", True
    End If

    ' Update the track time display
    mciSendString "status cd position", e, Len(e), 0
    Track = CInt(Mid$(e, 1, 2))
    Minute = CInt(Mid$(e, 4, 2))
    Second = CInt(Mid$(e, 7, 2))
    TimeWindow.Text = "[" & Format(Track, "00") & "] " & Format(Minute, "00") _
            & ":" & Format(Second, "00")
    TrackTime.Caption = "Track time: " & TrackLength(Track)
    TrackSelection.Text = TrackSelection.List(Track - 1)
    ' Check if CD is playing
    mciSendString "status cd mode", e, Len(e), 0
    Playing = (Mid$(e, 1, 7) = "playing")
Else
    Eject.Enabled = False
    ' Disable all the controls, clear the display
    If (CDLoad = True) Then
        Play.Enabled = False
        Pause.Enabled = False
        FastForward.Enabled = False
        Rewind.Enabled = False
        NextTrack.Enabled = False
        PreviousTrack.Enabled = False
        stpButton.Enabled = False
        CDLoad = False
        Playing = False
        TrackTime.Caption = ""
        TrackTime.Caption = ""
        TimeWindow.Text = ""
    End If
End If
End Sub
' Stop the CD
Private Sub stpButton_Click()
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
End Sub

Private Sub Timer1_Timer()
 FlashWindow hwnd, 1
Update

End Sub

Private Function fSetVolumeControl(ByVal hmixer As Long, _
    mxc As MIXERCONTROL, ByVal Volume As Long) As Boolean
'
' This function sets the value for a volume control.
'
Dim rc   As Long
Dim mxcd As MIXERCONTROLDETAILS
Dim vol  As MIXERCONTROLDETAILS_UNSIGNED

With mxcd
    .item = 0
    .dwControlID = mxc.dwControlID
    .cbStruct = Len(mxcd)
    .cbDetails = Len(vol)
End With
'
' Allocate a buffer for the control value buffer.
'
hmem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hmem)
mxcd.cChannels = 1
vol.dwValue = Volume
'
' Copy the data into the control value buffer.
'
Call CopyPtrFromStruct(mxcd.paDetails, vol, Len(vol))
'
' Set the control value.
'
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
Call GlobalFree(hmem)

If MMSYSERR_NOERROR = rc Then
    fSetVolumeControl = True
Else
    fSetVolumeControl = False
End If
End Function

Private Sub TrackSelection_Click()
lblTrackSelection.Visible = True
If (CDLoad) Then
        'Set TrackSelection value first
       
        If (Track <= TotalTracks) Then
            If (Playing) Then
                Command = "play cd from " & Val(TrackSelection.Text)
                SendMCIString Command, True
             Else
                Command = "seek cd to " & Val(TrackSelection.Text)
                SendMCIString Command, True
                SendMCIString "play cd", True
                Playing = True
            End If
        End If
        Else
        SendMCIString "seek cd to 1", True
    End If
    Update
End Sub

Private Sub Volume_Change()
lblVolume.Visible = True
Dim lVol As Long
lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub

Private Sub Volume_Scroll()
Dim lVol As Long

lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub

