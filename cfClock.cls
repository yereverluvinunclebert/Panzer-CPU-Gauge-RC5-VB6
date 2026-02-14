VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfGauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfGauge
' Author    : beededea
' Date      : 18/10/2024
' Purpose   : This is the class for the gauge form and all the elements upon it.
'---------------------------------------------------------------------------------------

'@IgnoreModule IntegerDataType, ModuleWithoutFolder
Option Explicit

'Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
'Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

' Declare gaugeForm as an RC form to be available throughout by calling fGauge as defined as this class in modMain

Public WithEvents gaugeForm As cWidgetForm
Attribute gaugeForm.VB_VarHelpID = -1

' public vars and objects

Public collPSDNonUIElements As cCollection

Public FX As Long   ' available externally, to allow the gaugeForm to be moved, need to add set/let for these public vars
Public FY As Long
Public FZ As Single

'Private mGlowing As Boolean

' do these need to be changed to properties?
'Private pauseElapsedTimeSecs As Long ' long means a max of 2,147,483,647 seconds, 24,855 days or 68 years
'Private restartElapsedTimeSecs As Long

'Private pausedSWSDeg As Long ' pausedSWSDeg = 0

' class private members for property assignment

'Private mMuteToggleEnabled As Boolean
'Private mTicking As Boolean
'Private mShowHelp As Boolean
Private mOpacity As String
Private pPSDWidth As Long
Private pPSDHeight As Long
'Private pBusyTimerRotateValue As Integer

' RC widgets

'Private WithEvents helpbottom As cWidgetBase

'local general vars used only within this class


'---------------------------------------------------------------------------------------
' Procedure : initialisePrivateVars
' Author    : beededea
' Date      : 12/02/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrivateVars()
   On Error GoTo initialisePrivateVars_Error

    pPSDWidth = 0
    pPSDHeight = 0

   On Error GoTo 0
   Exit Sub

initialisePrivateVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrivateVars of Class Module cfGauge"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Constructor (Initialise)
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   
    On Error GoTo Class_Initialize_Error
    
    Call initialisePrivateVars
    
    'pBusyTimerRotateValue = 1

    Set collPSDNonUIElements = New_c.Collection(False)
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfGauge"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : InitialiseImageWidgetsFromXML
' Author    : olaf schmidt and myself
' Date      : 31/07/2023
' Purpose   :  Using a previously created Cairo imageList, it creates a surface from each and every PSD layer
'              in the PSD file. The surfaces are populated from PNGs whose metric data are held in an XML file.
'              For each previously 'excluded' entry, it adds path X,Y and alpha properties to the excluded collection.
'              It creates an instance of the cwGaugeOverlay class and populates it with the excluded items that will be rendered in the overlay
'              The overlay comprises items that are non-clickable and will not generate events, ie. animated gauge hands, pendulums &c.
'
'              For all the interactive UI elements it creates RichClient widgets with corresponding keynames, locations and sizes as per the original PSD for each layer.
'              NOTE: The XML data is created using the Create Widget 2.0.js using Photoshop and the PSD found in the RES folder.
'              The XML file is renamed from .KON and placed in the RES folder.
'---------------------------------------------------------------------------------------
'
Public Sub InitialiseImageWidgetsFromXML()
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
    
    On Error GoTo InitialiseImageWidgetsFromXML_Error

    'create the Top-Level-Form
    Set gaugeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), gsWidgetName, True, 1, 1)
        gaugeForm.WidgetRoot.BackColor = -1 ' transparent
        
    ' code to read image XML data instead of directly from PSD, for RichClient5 only
    Call convertImageXMLToRcWidgets

    ' create a overlay here just for a single custom widget that contains just the animated elements.
    Set gaugeOverlay = gaugeForm.Widgets.Add(New cwGaugeOverlay, "Overlay", 0, 0, pPSDWidth, pPSDHeight)
    Set gaugeOverlay.cCollOverlayImages = collPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
        gaugeOverlay.cxPerc = 0.4940476 'define the relative center (within cwGaugeOverlay, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
        gaugeOverlay.cyPerc = 0.4727272
    
     ' set the various widgets as WithEvents Variables
'    Set tickbutton = gaugeForm.Widgets("tickbutton").Widget
'    Set helpbutton = gaugeForm.Widgets("helpbutton").Widget
'    Set startbutton = gaugeForm.Widgets("startbutton").Widget
'    Set stopbutton = gaugeForm.Widgets("stopbutton").Widget
'    Set middlebutton = gaugeForm.Widgets("middlebutton").Widget
'    Set lockbutton = gaugeForm.Widgets("lockbutton").Widget
'    Set prefsbutton = gaugeForm.Widgets("prefsbutton").Widget
'    Set surround = gaugeForm.Widgets("surround").Widget
    
    On Error GoTo 0
   Exit Sub

InitialiseImageWidgetsFromXML_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitialiseImageWidgetsFromXML of Class Module cfGauge"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_BubblingEvent
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events for all pre-defined PSD layer widgets
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_BubblingEvent(ByRef Sender As Object, ByRef EventName As String, ByRef p1 As Variant, ByRef p2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
    
    Dim stdMessage As String: stdMessage = vbNullString
    Dim thisMessage As String: thisMessage = vbNullString
    Dim thisSubject As String: thisSubject = vbNullString
    
    On Error GoTo gaugeForm_BubblingEvent_Error

    If gsIgnoreMouse = "1" Then Exit Sub
    
    If Not Sender.Widget.Enabled Then Exit Sub   ' Sender is disabled and bubbling events will be ignored
    
    ' set the balloon tooltips and perform other tasks onMouseMove, RC style tooltips are defined elsewhere
     If EventName = "W_MouseMove" Then
        Select Case LCase$(Sender.Widget.key)
            Case "lockbutton"
                If gsWidgetTooltips = "0" Then
                    thisMessage = "Click me to lock the widget in place. This lock can be released by clicking again on the lock button, or in the preferences. A locking sound will denote the button's operation."
                    thisSubject = " Lock Button Help"
                End If
            Case "prefsbutton"
                If gsWidgetTooltips = "0" Then
                    thisMessage = "Click me to open the preferences utility. The configuration of the widget is performed using the preference utility."
                    thisSubject = " Prefs Button Help"
                End If
            Case "tickbutton"
                If gsWidgetTooltips = "0" Then
                    thisMessage = "Click here to toggle the pointer movement from once per interval to a more fluid, continuous movement."
                    thisSubject = " Tick Button Help"
                End If
            Case "middlebutton"
                If gsWidgetTooltips = "0" Then
                    thisMessage = "This button switches on the mutli-core display. If your CPU has more than one core a bar will be shown, displaying the CPU for each core."
                    thisSubject = " Multi-Core Button Help"
                End If
            Case "stopbutton"
                Call stopbutton_mouseMove
            Case "startbutton"
                Call startbutton_mouseMove
            Case "helpbutton"
                If gsWidgetTooltips = "0" Then
                    thisMessage = "This button opens the one-page help file for this widget. There is a further detailed CHM help file available from the right-click menu that contains all the information that you need to run this widget."
                    thisSubject = " Help Button Help"
                End If
            Case Else
                If gsWidgetTooltips = "0" Then
                    stdMessage = "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize. "
                    stdMessage = stdMessage & vbCrLf & "You can turn off the balloon tooltips in the preferences."
                    thisMessage = sCPUDetailStore & vbCrLf & stdMessage
                    thisSubject = " Information"
                End If
        End Select
        
        ' now set the tooltip for the whole gauge using the derived message and subject
        If gsWidgetTooltips = "0" Then CreateToolTip gaugeForm.hwnd, thisMessage, _
            TTIconInfo, gsWidgetName & thisSubject, , , , True
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(Sender.Widget.key)
            Case "lockbutton"
                Call lockbutton_mouseDown(Sender)
            Case "prefsbutton"
                Call prefsbutton_mouseDown(Sender)
            Case "tickbutton"
                Call tickbutton_mouseDown(Sender)
            Case "middlebutton"
                 Call middlebutton_mouseDown(Sender)
            Case "stopbutton"
                Call stopbutton_mouseDown(Sender)
            Case Else
                Debug.Print Sender.Widget.key
        End Select
    End If
    
    If EventName = "W_MouseUp" Then
        Select Case LCase$(Sender.Widget.key)
            Case "prefsbutton"
                Call prefsbutton_mouseUp(Sender)
            Case "middlebutton"
                Call middlebutton_mouseUp(Sender)
            Case "stopbutton"
                Call stopbutton_mouseUp(Sender)
            Case Else
                Debug.Print Sender.Widget.key
        End Select
    End If
    
    If EventName = "W_Click" Then
        Select Case LCase$(Sender.Widget.key)
            Case "helpbutton"
                 Call helpbutton_click
            Case "startbutton"
                 Call startbutton_click
            Case Else
                Debug.Print Sender.Widget.key
        End Select
    End If
            
   On Error GoTo 0
   Exit Sub

gaugeForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_BubblingEvent of Class Module cfGauge"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : startbutton_mouseMove
' Author    : beededea
' Date      : 07/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub startbutton_mouseMove()
    On Error GoTo startbutton_mouseMove_Error

    If gsWidgetTooltips = "0" Then
            CreateToolTip gaugeForm.hwnd, "Clock mode, this button restarts all stopped clock functions for this widget, hand movement &c .", _
                TTIconInfo, gsWidgetName & " Start Button Help", , , , True
    End If

    On Error GoTo 0
    Exit Sub

startbutton_mouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startbutton_mouseMove of Class Module cfGauge"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : stopbutton_mouseMove
' Author    : beededea
' Date      : 07/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_mouseMove()
    On Error GoTo stopbutton_mouseMove_Error

    If gsWidgetTooltips = "0" Then
            CreateToolTip gaugeForm.hwnd, "Clock mode, this button stops all animated clock functions for this widget, hand movements &c .", _
                TTIconInfo, gsWidgetName & " Stop Button Help", , , , True
    End If

    On Error GoTo 0
    Exit Sub

stopbutton_mouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseMove of Class Module cfGauge"
End Sub

                


'---------------------------------------------------------------------------------------
' Procedure : helpbutton_click
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helpbutton_click()
   On Error GoTo helpbutton_click_Error

   Call helpSplash

   On Error GoTo 0
   Exit Sub

helpbutton_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbutton_click of Class Module cfGauge"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : startbutton_click
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub startbutton_click()

    On Error GoTo startbutton_click_Error
    
    Dim fileToPlay As String: fileToPlay = vbNullString

    Call TurnFunctionsOn

    fileToPlay = "ting.wav"
    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If


   On Error GoTo 0
   Exit Sub

startbutton_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startbutton_click of Class Module cfGauge"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : stopbutton_mouseUp
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_mouseUp(ByRef Sender As Object)
     Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo stopbutton_mouseUp_Error

    Sender.Widget.Alpha = Val(gsOpacity) / 100
    Sender.Widget.Refresh
           
    Call SwitchOff
    
    fileToPlay = "stop.wav"
    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Sub

stopbutton_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseUp of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : stopbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_mouseDown(ByRef Sender As Object)
   
    On Error GoTo stopbutton_mouseDown_Error

    Sender.Widget.Alpha = 0
    Sender.Widget.Refresh
           
   On Error GoTo 0
   Exit Sub

stopbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseDown of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : middlebutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : toggles the display of the multi-core VB6 form
'---------------------------------------------------------------------------------------
'
Private Sub middlebutton_mouseDown(ByRef Sender As Object)
    On Error GoTo middlebutton_mouseDown_Error

    Sender.Widget.Alpha = 0
    Sender.Widget.Refresh
    
    If gsMultiCoreEnable = "0" Then
        gsMultiCoreEnable = "1"
        frmMultiCore.Show
        Call startAllCpuTimers
    Else
        gsMultiCoreEnable = "0"
        frmMultiCore.Hide
        Call stopAllCpuTimers
    End If
    
    widgetPrefs.chkMultiCoreEnable.Value = Val(gsMultiCoreEnable)
    
   On Error GoTo 0
   Exit Sub

middlebutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure middlebutton_mouseDown of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : middlebutton_mouseUp
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub middlebutton_mouseUp(ByRef Sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
    On Error GoTo middlebutton_mouseUp_Error

    Sender.Widget.Alpha = Val(gsOpacity) / 100
    Sender.Widget.Refresh
        
    fileToPlay = "till.wav"
    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    
   On Error GoTo 0
   Exit Sub

middlebutton_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure middlebutton_mouseUp of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : prefsbutton_mouseUp
' Author    : beededea
' Date      : 07/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub prefsbutton_mouseUp(ByRef Sender As Object)
    On Error GoTo prefsbutton_mouseUp_Error

    Sender.Widget.Alpha = Val(gsOpacity) / 100
    Sender.Widget.Refresh

    On Error GoTo 0
    Exit Sub

prefsbutton_mouseUp_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prefsbutton_mouseUp of Class Module cfGauge"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lockbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockbutton_mouseDown(ByRef Sender As Object)
   On Error GoTo lockbutton_mouseDown_Error
    
    If gaugeOverlay.Locked = False Then
        Sender.Widget.Alpha = 0
    Else
        Sender.Widget.Alpha = Val(gsOpacity) / 100
    End If
    Sender.Widget.Refresh ' this removes the delay in the lockButton depressing
    Call toggleWidgetLock
    
   On Error GoTo 0
   Exit Sub

lockbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockbutton_mouseDown of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : prefsbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub prefsbutton_mouseDown(ByRef Sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo prefsbutton_mouseDown_Error

    Sender.Widget.Alpha = 0
    'Sender.Widget.Visible = False
    Sender.Widget.Refresh ' this removes the delay in the Button depressing
    fileToPlay = "winding.wav"
    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
   'MsgBox "prefsbutton_mouseDown"
   Call makeProgramPreferencesAvailable

   On Error GoTo 0
   Exit Sub

prefsbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prefsbutton_mouseDown of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : tickbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tickbutton_mouseDown(ByRef Sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo tickbutton_mouseDown_Error
        
    If gaugeOverlay.PointerAnimate = True Then
        gaugeOverlay.PointerAnimate = False
        Sender.Widget.Alpha = Val(gsOpacity) / 100
        gsPointerAnimate = "0"
    Else
        gaugeOverlay.PointerAnimate = True
        Sender.Widget.Alpha = 0 ' sender is the referring control
        gsPointerAnimate = "1"
        
    End If
    Sender.Widget.Refresh ' this removes the delay in the Button depressing
    
    fileToPlay = "lock.wav"
    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    sPutINISetting "Software\PzCPUGauge", "pointerAnimate", gsPointerAnimate, gsSettingsFile
    widgetPrefs.cmbTickSwitchPref.ListIndex = Val(gsPointerAnimate)
    
   On Error GoTo 0
   Exit Sub

tickbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tickbutton_mouseDown of Class Module cfGauge"

End Sub


    



'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : If a double-click command has been selected in the prefs then this will carry it out
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo gaugeForm_DblClick_Error

    '    If gsIgnoreMouse = "1" Then Exit Sub

    If gsWidgetFunctions = "0" Or gsIgnoreMouse = "1" Then Exit Sub
    
    If LTrim$(gsDblClickCommand) = vbNullString Then Exit Sub
    
    thisCommand = gsDblClickCommand
        
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
    If gbSHIFT_1 = True Then
        gbSHIFT_1 = False
        Call ShellExecute(gaugeForm.hwnd, "Open", gsOpenFile, vbNullString, App.Path, 1)
    Else
        Call ShellExecute(gaugeForm.hwnd, "runas", thisCommand, vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

gaugeForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_DblClick of Class Module cfGauge"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses and key downs for the main toggles H,A,S,M,P,T,D,W,B & 1-5
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo gaugeForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

    On Error GoTo 0
    Exit Sub

gaugeForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_KeyDown of Class Module cfGauge"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :    if this is a multi monitor system then:
'                we note the monitor ID at gaugeForm form_load and store as the gaugeFormMonitorID, for widgetPrefs widgetPrefsMonitorID
'                on gaugeFromMouseUp (drag complete) we sample the monitor ID again -only works, for the prefs form we have to have a 200ms timer that tests for form .left and .top x,y change (no MOVED event and mouse UP does not work on the form itself when the titlebar dragged)
'                if the monitor has changed then sample the physical monitor resolution
'                if the resolution is different then calculate new size proportion
'                alter the size of the form to correspond to the monitor
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)

    On Error GoTo gaugeForm_MouseUp_Error
    
    gblGaugeCPUTimersOFF = False
    
    ' on the mouseDown event the sampling timers are stopped as extracting the CPU details causes a reduction in responsiveness to user input, this restores the sampling timers
    ' on a specified interval that will be corrected later by the timer itself when it next runs.
    
    gaugeOverlay.tmrSampler.Interval = 3000
    
    Call startAllCpuTimers
    
    frmTimer.tmrScreenResolution.Interval = 4500

    ' save the form x,y position whenever the globe/form is dragged
    Call saveMainRCFormPosition

    Call resizeLocateRCFormByMoveToNewMonitor
    
    On Error GoTo 0
    Exit Sub

gaugeForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseUp of Class Module cfGauge"
            Resume Next
          End If
    End With
End Sub




'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseMove
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : move the whole form on a mouseDown and drag
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = y 'just store the offsets when no button is down
  
   On Error GoTo gaugeForm_MouseMove_Error
    
    If gaugeOverlay.Locked = True Then Exit Sub
    If gsIgnoreMouse = "1" Then Exit Sub
    If gbMenuOccurred = True Then
        gbMenuOccurred = False
        Exit Sub
    End If

    If Button = vbLeftButton And Not gaugeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (x - x0) * gaugeForm.WidgetRoot.Zoom + gaugeForm.Left
      FY = (y - y0) * gaugeForm.WidgetRoot.Zoom + gaugeForm.Top
      gaugeForm.Move FX, FY
      
      Debug.Print "FX " & FX
      Debug.Print "FY " & FY
      
    End If


   On Error GoTo 0
   Exit Sub

gaugeForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseMove of Class Module cfGauge"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseWheel
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : Handles rotation of the mousewheel, with CTRL key pressed, making the widget smaller or larger. By default, rotate forward positive = smaller
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)

    Dim sizingDirection As Long: sizingDirection = 0
    Dim wheelRotation As Long: wheelRotation = 0
   
    On Error GoTo gaugeForm_MouseWheel_Error
    
    'sizingDirection = Abs(wheelRotation)
    wheelRotation = Rotation / 2
    
    ' ctrl pressed at the same time so we are resizing the whole thing
    If MouseKeys = 8 Then
    
        If gsScrollWheelDirection = "1" Then
            If wheelRotation > 0 Then
                sizingDirection = 0 - Abs(wheelRotation)
            Else
                sizingDirection = Abs(wheelRotation)
            End If
        Else
            sizingDirection = Abs(wheelRotation)
        End If
        
        ' adjust the slider in the widget Prefs to centralise resizing and prevent re-entrancy, it calls AdjustZoom
        widgetPrefs.GaugeSize = (FZ + sizingDirection / 2400) 'change the Zoom only, when the Ctrl-Key is down
        
        ' saves the current ratios for the RC form
        Call saveRCFormCurrentSizeRatios
        
    End If

    On Error GoTo 0
    Exit Sub

gaugeForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseWheel of Class Module cfGauge"
End Sub


'public sub

'---------------------------------------------------------------------------------------
' Procedure : AdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Adjust the size of the whole gauge, anchoring the gauge top left, storing the values for later use
'---------------------------------------------------------------------------------------
'
Public Sub AdjustZoom(ByVal NewZoom As Single)
    On Error GoTo AdjustZoom_Error

    FZ = NewZoom
    If FZ < 0.05 Then FZ = 0.05 Else If FZ > 3 Then FZ = 3 'keep FZ within a sane ZoomInterval
    gaugeForm.WidgetRoot.Zoom = FZ '* gaugeForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
    gaugeForm.Move gaugeForm.Left, gaugeForm.Top, pPSDWidth * gaugeForm.WidgetRoot.Zoom, pPSDHeight * gaugeForm.WidgetRoot.Zoom
    
    ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
    gsWidgetSize = CStr(NewZoom * 100)
    
    If widgetPrefs.IsLoaded Then widgetPrefs.sliGaugeSize.Value = Val(gsWidgetSize) ' ensure that the prefs size slider matches when zoom is modified without using the slider
    
    If gWidgetMonitorStruct.IsPrimary = True Then
        gsWidgetPrimaryHeightRatio = CStr(NewZoom)
    Else
        gsWidgetSecondaryHeightRatio = CStr(NewZoom)
    End If

   On Error GoTo 0
   Exit Sub

AdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdjustZoom of Class Module cfGauge"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   : A click on the gauge form, test screen resolution and handle right mouse click for the menu
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo gaugeForm_MouseDown_Error
   
    ' s global variable flag to tell all animate and CPU timers to exit before running
    gblGaugeCPUTimersOFF = True
   
    ' on the mouseDown event the sampling timers are stopped as extracting the CPU details causes a reduction in responsiveness to user input, these timers are restarted during the
    ' gaugeForm_MouseUp event when the user lets go of the gauge.
    
    Call stopAllCpuTimers
    
    ' here run the timer to restart the sampling timers for both the CPU and multicore forms
    frmTimer.tmrRestartSampling.Enabled = True

    'essential code to test for monitor resolution before and after a mouse-down pick up and drag to another monitor
    frmTimer.tmrScreenResolution.Enabled = False
    frmTimer.tmrScreenResolution.Interval = 200
    frmTimer.tmrScreenResolution.Enabled = True
        
    If Button = vbRightButton Then
        gbMenuOccurred = True
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If gsIgnoreMouse = "1" Then Exit Sub
        ' so we pass it through
    End If

   On Error GoTo 0
   Exit Sub

gaugeForm_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseDown of Class Module cfGauge"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsVisible
' Author    : beededea
' Date      : 08/05/2023
'
' Purpose   : Useful property to determine by value as to whether this form is loaded,
'             allowing external checks to the form to determine whether it is loaded,
'             WITHOUT activating the form automatically.
'---------------------------------------------------------------------------------------
'
Public Property Get IsVisible() As Boolean
    On Error GoTo IsVisible_Error

    If gaugeForm.WindowState = vbNormal Then
        IsVisible = gaugeForm.Visible
    Else
        IsVisible = False
    End If

    On Error GoTo 0
    Exit Property

IsVisible_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Class Module cfGauge"
            Resume Next
          End If
    End With
End Property



'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the opacity of the elements within this class
'---------------------------------------------------------------------------------------
'
Public Property Get Opacity() As String
   On Error GoTo OpacityGet_Error

   Opacity = mOpacity

   On Error GoTo 0
   Exit Property

OpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cfGauge"
End Property
'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the opacity of the elements within this class
'---------------------------------------------------------------------------------------
'
Public Property Let Opacity(ByVal newValue As String)
   On Error GoTo OpacityLet_Error

   If mOpacity <> newValue Then mOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

OpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cfGauge"
End Property







' ----------------------------------------------------------------
' Procedure Name: convertImageXMLToRcWidgets
' Purpose: code to read XML image data instead of reading it directly from PSD which is RC6 only, an alternative for RichClient5 use.
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 28/08/2025
' ----------------------------------------------------------------
Private Sub convertImageXMLToRcWidgets()

    Dim W As cWidgetBase

    'Dim answer As VbMsgBoxResult
    Dim answerMsg  As String: answerMsg = vbNullString

    Dim num_results As Integer: num_results = 0
    Dim windowWidth As String: windowWidth = vbNullString
    Dim windowHeight As String: windowHeight = vbNullString
    Dim sSrc  As String: sSrc = vbNullString
    Dim sName As String: sName = vbNullString
    Dim sWidth As String: sWidth = vbNullString
    Dim sHeight As String: sHeight = vbNullString
'    Dim sHOffset As String: sHOffset = vbNullString
'    Dim sVOffset As String: sVOffset = vbNullString
'    Dim sOpacity As String: sOpacity = vbNullString
    Dim Width As Long: Width = 0
    Dim height As Long: height = 0
    Dim hOffset As Long: hOffset = 0
    Dim vOffset As Long: vOffset = 0
    Dim Opacity As Integer: Opacity = 0
    Dim someOpacity As Double: someOpacity = 0
    Dim xmlFileToLoad As String: xmlFileToLoad = vbNullString
    Dim pngFileToLoad As String: pngFileToLoad = vbNullString
    
    Dim nodeList As MSXML2.IXMLDOMNodeList
    Dim objxmldoc As MSXML2.DOMDocument
   
    Set objxmldoc = New MSXML2.DOMDocument
    
    Dim node As MSXML2.IXMLDOMNode
    Dim MainNode As MSXML2.IXMLDOMNode
'    Dim ImageNode As MSXML2.IXMLDOMNode
'    Dim ImageNodes As MSXML2.IXMLDOMNodeList
    
    On Error GoTo convertImageXMLToRcWidgets_Error
    
    someOpacity = Val(mOpacity) / 100
    xmlFileToLoad = App.Path & "\RES\Panzer-CPU-gauge-VB6.xml"
    
    If fFExists(xmlFileToLoad) Then
        objxmldoc.Load xmlFileToLoad
    Else
        MsgBox "The XML file that contains the image data is missing " & xmlFileToLoad
    End If
    
    ' obtain the overall widget width and height
    Set MainNode = objxmldoc.selectSingleNode("widget/window")
            
    windowWidth = MainNode.selectSingleNode("@width").Text
    windowHeight = MainNode.selectSingleNode("@height").Text

    pPSDWidth = CLng(windowWidth)
    pPSDHeight = CLng(windowHeight)
    
    ' get the image values from the XML data, the num results should be non-zero
    Set nodeList = objxmldoc.selectNodes("widget/window/image")
    num_results = nodeList.Length
    
    ' no results found, go on as normal using the sampling interval
    If num_results = 0 Then
        answerMsg = "1. There is a problem with the XML data file that describes the image, seems to contain no valid data"
        msgBoxA answerMsg, vbOKOnly + vbExclamation, "XML Warning", True, "convertImageXMLToRcWidgetsPollingWarning"
        Exit Sub ' Return
    End If

    If Not nodeList Is Nothing Then
         For Each node In nodeList
         
            sSrc = node.selectSingleNode("@src").Text
            sSrc = Replace(sSrc, "/", "\")
            sName = node.selectSingleNode("@name").Text
            sWidth = node.selectSingleNode("@width").Text
            Width = CLng(Left$(sWidth, (InStr(sWidth, " px") - 1)))
            
            sHeight = node.selectSingleNode("@height").Text
            height = CLng(Left$(sHeight, (InStr(sHeight, " px") - 1)))
            
            hOffset = CLng(node.selectSingleNode("@hOffset").Text)
            vOffset = CLng(node.selectSingleNode("@vOffset").Text)
            Opacity = CInt(node.selectSingleNode("@opacity").Text) / 2.55
            
            If Opacity = 100 Then  ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
            
               'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                pngFileToLoad = App.Path & "\RES\" & sSrc
                If fFExists(pngFileToLoad) Then
                    Cairo.ImageList.AddSurface sName, Cairo.CreateSurface(Width, height, PSSurface, pngFileToLoad)
                Else
                    MsgBox "Error, this PNG resource file seems to be missing " & pngFileToLoad
                End If
                
                ' check if each layer is in the layer exclude list, if it IS then we add it to a collection for non UI elements (ie. do not create Widgets)
                If collPSDNonUIElements.Exists(sName) Then
                  
                     'we add layer info. (used later in cwGaugeOverlay) to the excluded layers that will form the overlay.
                     collPSDNonUIElements(sName) = Array(hOffset, vOffset, someOpacity)  'here we update the so far empty slots with the PSD-offsets
                Else

                     'create a widget instance for all layers in the PSD, excluding any layers entered into the exclude-list
                     Set W = gaugeForm.Widgets.Add(New cwAlphaImg, LCase$(sName), hOffset, vOffset, Width, height).Widget

                     W.ImageKey = W.key 'W.Key equals ImageList-Key, set above - and sName at this point ... set it also as the ImageKey of our new created Widget
                     W.Alpha = 1

                     ' note: the clickable layers characteristics are set in adjustMainControls
                     ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                     W.HoverColor = -1
                     If gsWidgetTooltips = "1" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                     W.MousePointer = IDC_SIZEALL

                 End If
            End If
          Next node
    End If
   
    'Cleanup
    Set nodeList = Nothing
    Set MainNode = Nothing
    
    On Error GoTo 0
    Exit Sub

convertImageXMLToRcWidgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure convertImageXMLToRcWidgets, line " & Erl & "."

End Sub
