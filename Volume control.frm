VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form VolumeControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Volume Control"
   ClientHeight    =   7170
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3150
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Volume control.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.VScrollBar Volume 
      Height          =   6855
      LargeChange     =   2
      Left            =   120
      MousePointer    =   4  'Icon
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   3720
   End
End
Attribute VB_Name = "VolumeControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Command As String          ' string to hold mci command strings
Dim hmixer As Long             ' mixer handle
Dim volCtrl As MIXERCONTROL    ' Waveout volume control.


Private Sub Form_Load()
Dim rc  As Long
Dim OK As Boolean
' Open the mixer with deviceID 0.
'
rc = mixerOpen(hmixer, 0, 0, 0, 0)
If MMSYSERR_NOERROR <> rc Then
    MsgBox "Could not open the mixer.", vbCritical, "Volume Control"
    Exit Sub
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

Private Sub Volume_Change()
Dim lVol As Long
lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub
