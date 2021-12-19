Attribute VB_Name = "VolumeControl_Module"
Option Explicit

'API Calls

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'
'-----------------------------------------------
' Wave file related.
'-----------------------------------------------
'
Public hmem As Long

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_PURGE = &H40
Public Const SND_FILENAME = &H20000

Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
               (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CONTROLTYPE_FADER = _
               (MIXERCONTROL_CT_CLASS_FADER Or _
               MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = _
               (MIXERCONTROL_CONTROLTYPE_FADER + 1)

Public Type MIXERCONTROLDETAILS
    cbStruct    As Long
    dwControlID As Long
    cChannels   As Long
    item        As Long
    cbDetails   As Long
    paDetails   As Long
End Type

Public Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long
End Type

Public Type MIXERCONTROL
    cbStruct       As Long
    dwControlID    As Long
    dwControlType  As Long
    fdwControl     As Long
    cMultipleItems As Long
    szShortName    As String * MIXER_SHORT_NAME_CHARS
    szName         As String * MIXER_LONG_NAME_CHARS
    lMinimum       As Long
    lMaximum       As Long
    reserved(10)   As Long
End Type

Public Type MIXERLINECONTROLS
    cbStruct  As Long
    dwLineID  As Long
    dwControl As Long
    cControls As Long
    cbmxctrl  As Long
    pamxctrl  As Long
End Type

Public Type MIXERLINE
    cbStruct        As Long
    dwDestination   As Long
    dwSource        As Long
    dwLineID        As Long
    fdwLine         As Long
    dwUser          As Long
    dwComponentType As Long
    cChannels       As Long
    cConnections    As Long
    cControls       As Long
    szShortName     As String * MIXER_SHORT_NAME_CHARS
    szName          As String * MIXER_LONG_NAME_CHARS
    dwType          As Long
    dwDeviceID      As Long
    wMid            As Integer
    wPid            As Integer
    vDriverVersion  As Long
    szPname         As String * MAXPNAMELEN
End Type
'
'Allocates the specified number of bytes from the heap.
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
'
'Locks a global memory object and returns a pointer to the
' first byte of the object's memory block.  The memory block
' associated with a locked object cannot be moved or discarded.
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
'
'Frees the specified global memory object and invalidates its handle.
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
'
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal ptr As Long, struct As Any, ByVal cb As Long)

Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" _
    (struct As Any, ByVal ptr As Long, ByVal cb As Long)
'
'Opens a specified mixer device and ensures that the device
' will not be removed until the application closes the handle.
Declare Function mixerOpen Lib "winmm.dll" _
    (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, _
    ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
'
'Sets properties of a single control associated with an audio line.
Declare Function mixerSetControlDetails Lib "winmm.dll" _
    (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long) As Long
'
'Retrieves information about a specific line of a mixer device.
Declare Function mixerGetLineInfo Lib "winmm.dll" _
    Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, _
    pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
'
'Retrieves one or more controls associated with an audio line.
Declare Function mixerGetLineControls Lib "winmm.dll" _
    Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, _
    pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
'
'Play a wave file.
Public Declare Function PlaySound Lib "winmm.dll" _
    Alias "PlaySoundA" (ByVal lpszName As String, _
    ByVal hModule As Long, ByVal dwFlags As Long) As Long





Public Function fGetVolumeControl(ByVal hmixer As Long, _
        ByVal componentType As Long, ByVal ctrlType As Long, _
        ByRef mxc As MIXERCONTROL) As Boolean
'
' This function attempts to obtain a mixer control.
'
Dim mxlc As MIXERLINECONTROLS
Dim mxl  As MIXERLINE
Dim hmem As Long
Dim rc   As Long

mxl.cbStruct = Len(mxl)
mxl.dwComponentType = componentType
'
' Get a line corresponding to the component type.
'
rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
If MMSYSERR_NOERROR = rc Then
    With mxlc
        .cbStruct = Len(mxlc)
        .dwLineID = mxl.dwLineID
        .dwControl = ctrlType
        .cControls = 1
        .cbmxctrl = Len(mxc)
    End With
    '
    ' Allocate a buffer for the control.
    '
    hmem = GlobalAlloc(&H40, Len(mxc))
    mxlc.pamxctrl = GlobalLock(hmem)
    mxc.cbStruct = Len(mxc)
    '
    ' Get the control.
    '
    rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
    If MMSYSERR_NOERROR = rc Then
        fGetVolumeControl = True
        '
        ' Copy the control into the destination structure.
        '
        Call CopyStructFromPtr(mxc, mxlc.pamxctrl, Len(mxc))
    Else
        fGetVolumeControl = False
    End If
    Call GlobalFree(hmem)
    Exit Function
End If
fGetVolumeControl = False
End Function











