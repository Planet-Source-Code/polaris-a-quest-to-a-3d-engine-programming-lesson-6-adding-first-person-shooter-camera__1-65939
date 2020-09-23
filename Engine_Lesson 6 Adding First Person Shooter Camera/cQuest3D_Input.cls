VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuest3D_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'19 March 2003

'========================================================
'
'  JOYSTICK functions and routines
'
'
'=======================================================
Option Explicit

'This project needs 2 Buttons
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursor Lib "user32" () As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Dim aKeys(255) As String

Public Enum QUEST3D_KEY_STATE
    QUEST3D_KEYUP = 0

    QUEST3D_KEYDOWN = 1
End Enum


Enum QUEST3D_MOUSE_BUTTON

    QUEST3D_MOUSE_LEFT_Button = 0
    QUEST3D_MOUSE_RIGHT_Button = 1
    QUEST3D_MOUSE_MIDDLE_Button = 2

End Enum

Public Enum QUEST3D_INPUTDEVICEICETYPE
    QUEST3D_INPUTDEVICE_CLASS_ALL = 0
    QUEST3D_INPUTDEVICE_CLASS_DEVICE = 1
    QUEST3D_INPUTDEVICE_CLASS_GAMECTRL = 4
    QUEST3D_INPUTDEVICE_CLASS_KEYBOARD = 3
    QUEST3D_INPUTDEVICE_CLASS_POINTER = 2

    QUEST3D_INPUTDEVICE_TYPE_1STPERSON = 24     '(&H18)
    QUEST3D_INPUTDEVICE_TYPE_DEVICE = 17        '(&H11)
    QUEST3D_INPUTDEVICE_TYPE_DEVICECTRL = 25    '(&H19)
    QUEST3D_INPUTDEVICE_TYPE_DRIVING = 22       '(&H16)
    QUEST3D_INPUTDEVICE_TYPE_FLIGHT = 23        '(&H17)
    QUEST3D_INPUTDEVICE_TYPE_GAMEPAD = 21       '(&H15)
    QUEST3D_INPUTDEVICE_TYPE_JOYSTICK = 20      '(&H14)
    QUEST3D_INPUTDEVICE_TYPE_KEYBOARD = 19      '(&H13)
    QUEST3D_INPUTDEVICE_TYPE_MOUSE = 18         '(&H12)
    QUEST3D_INPUTDEVICE_TYPE_REMOTE = 27        '(&H1B)
    QUEST3D_INPUTDEVICE_TYPE_SCREENPOINTER = 26 '(&H1A)
    QUEST3D_INPUTDEVICE_TYPE_SUPPLEMENTAL = 28  '(&H1C)
End Enum

Public Enum QUEST3D_GAMEPAD_STATE
    QUEST3D_GAMEPAD_UP ' =&h255
    QUEST3D_GAMEPAD_DOWN ' =&h255 Or 2
    QUEST3D_GAMEPAD_LEFT ' =&h255 Or 4
    QUEST3D_GAMEPAD_RIGHT ' =&h255 Or 8
    QUEST3D_GAMEPAD_DIAG_UPRIGHT ' =&h255 Or 16
    QUEST3D_GAMEPAD_DIAG_UPLEFT ' =&h255 Or 32
    QUEST3D_GAMEPAD_DIAG_DOWNRIGHT ' =&h255 Or 64
    QUEST3D_GAMEPAD_DIAG_DOWNLEFT ' =&h255 Or 128
    QUEST3D_GAMEPAD_BUTTON0 = 48  '(&H30)
    QUEST3D_GAMEPAD_BUTTON1 = 49  '(&H31)
    QUEST3D_GAMEPAD_BUTTON2 = 50  '(&H32)
    QUEST3D_GAMEPAD_BUTTON3 = 51  '(&H33)
    QUEST3D_GAMEPAD_BUTTON4 = 52  '(&H34)
    QUEST3D_GAMEPAD_BUTTON5 = 53  '(&H35)
    QUEST3D_GAMEPAD_BUTTON6 = 54  '(&H36)
    QUEST3D_GAMEPAD_BUTTON7 = 55  '(&H37)
    QUEST3D_GAMEPAD_BUTTON8 = 56  '(&H38)
    QUEST3D_GAMEPAD_BUTTON9 = 57  '(&H39)
    QUEST3D_GAMEPAD_BUTTON10 = 58 '(&H3A)
    QUEST3D_GAMEPAD_BUTTON11 = 59 '(&H3B)
    QUEST3D_GAMEPAD_BUTTON12 = 60 '(&H3C)
    QUEST3D_GAMEPAD_BUTTON13 = 61 '(&H3D)
    QUEST3D_GAMEPAD_BUTTON14 = 62 '(&H3E)
    QUEST3D_GAMEPAD_BUTTON15 = 63 '(&H3F)
    QUEST3D_GAMEPAD_BUTTON16 = 64 '(&H40)
    QUEST3D_GAMEPAD_BUTTON17 = 65 '(&H41)
    QUEST3D_GAMEPAD_BUTTON18 = 66 '(&H42)
    QUEST3D_GAMEPAD_BUTTON19 = 67 '(&H43)
    QUEST3D_GAMEPAD_BUTTON20 = 68 '(&H44)
    QUEST3D_GAMEPAD_BUTTON21 = 69 '(&H45)
    QUEST3D_GAMEPAD_BUTTON22 = 70 '(&H46)
    QUEST3D_GAMEPAD_BUTTON23 = 71 '(&H47)
    QUEST3D_GAMEPAD_BUTTON24 = 72 '(&H48)
    QUEST3D_GAMEPAD_BUTTON25 = 73 '(&H49)
    QUEST3D_GAMEPAD_BUTTON26 = 74 '(&H4A)
    QUEST3D_GAMEPAD_BUTTON27 = 75 '(&H4B)
    QUEST3D_GAMEPAD_BUTTON28 = 76 '(&H4C)
    QUEST3D_GAMEPAD_BUTTON29 = 77 '(&H4D)
    QUEST3D_GAMEPAD_BUTTON30 = 78 '(&H4E)
    QUEST3D_GAMEPAD_BUTTON31 = 79 '(&H4F)
End Enum

Private Sub Class_Initialize()

    aKeys(&H1) = "DIK_ESCAPE"
    aKeys(&H2) = "DIK_1"
    aKeys(&H3) = "DIK_2"
    aKeys(&H4) = "DIK_3"
    aKeys(&H5) = "DIK_4"
    aKeys(&H6) = "DIK_5"
    aKeys(&H7) = "DIK_6"
    aKeys(&H8) = "DIK_7"
    aKeys(&H9) = "DIK_8"
    aKeys(&HA) = "DIK_9"
    aKeys(&HB) = "DIK_0"
    aKeys(&HC) = "DIK_MINUS"             ' - on main keyboard
    aKeys(&HD) = "DIK_EQUALS"
    aKeys(&HE) = "DIK_BACK"                  ' backspace
    aKeys(&HF) = "DIK_TAB"
    aKeys(&H10) = "DIK_Q"
    aKeys(&H11) = "DIK_W"
    aKeys(&H12) = "DIK_E"
    aKeys(&H13) = "DIK_R"
    aKeys(&H14) = "DIK_T"
    aKeys(&H15) = "DIK_Y"
    aKeys(&H16) = "DIK_U"
    aKeys(&H17) = "DIK_I"
    aKeys(&H18) = "DIK_O"
    aKeys(&H19) = "DIK_P"
    aKeys(&H1A) = "DIK_LBRACKET"
    aKeys(&H1B) = "DIK_RBRACKET"
    aKeys(&H1C) = "DIK_RETURN" ' Enter on main keyboard
    aKeys(&H1D) = "DIK_LCONTROL"
    aKeys(&H1E) = "DIK_A"
    aKeys(&H1F) = "DIK_S"
    aKeys(&H20) = "DIK_D"
    aKeys(&H21) = "DIK_F"
    aKeys(&H22) = "DIK_G"
    aKeys(&H23) = "DIK_H"
    aKeys(&H24) = "DIK_J"
    aKeys(&H25) = "DIK_K"
    aKeys(&H26) = "DIK_L"
    aKeys(&H27) = "DIK_SEMICOLON"
    aKeys(&H28) = "DIK_APOSTROPHE"
    aKeys(&H29) = "DIK_GRAVE" ' accent grave
    aKeys(&H2A) = "DIK_LSHIFT"
    aKeys(&H2B) = "DIK_BACKSLASH"
    aKeys(&H2C) = "DIK_Z"
    aKeys(&H2D) = "DIK_X"
    aKeys(&H2E) = "DIK_C"
    aKeys(&H2F) = "DIK_V"
    aKeys(&H30) = "DIK_B"
    aKeys(&H31) = "DIK_N"
    aKeys(&H32) = "DIK_M"
    aKeys(&H33) = "DIK_COMMA"
    aKeys(&H34) = "DIK_PERIOD" ' . on main keyboard
    aKeys(&H35) = "DIK_SLASH" ' / on main keyboard
    aKeys(&H36) = "DIK_RSHIFT"
    aKeys(&H37) = "DIK_MULTIPLY" ' * on numeric keypad
    aKeys(&H38) = "DIK_LMENU" ' left Alt
    aKeys(&H39) = "DIK_SPACE"
    aKeys(&H3A) = "DIK_CAPITAL"
    aKeys(&H3B) = "DIK_F1"
    aKeys(&H3C) = "DIK_F2"
    aKeys(&H3D) = "DIK_F3"
    aKeys(&H3E) = "DIK_F4"
    aKeys(&H3F) = "DIK_F5"
    aKeys(&H40) = "DIK_F6"
    aKeys(&H41) = "DIK_F7"
    aKeys(&H42) = "DIK_F8"
    aKeys(&H43) = "DIK_F9"
    aKeys(&H44) = "DIK_F10"
    aKeys(&H45) = "DIK_NUMLOCK"
    aKeys(&H46) = "DIK_SCROLL" ' Scroll Lock
    aKeys(&H47) = "DIK_NUMPAD7"
    aKeys(&H48) = "DIK_NUMPAD8"
    aKeys(&H49) = "DIK_NUMPAD9"
    aKeys(&H4A) = "DIK_SUBTRACT" ' - on numeric keypad
    aKeys(&H4B) = "DIK_NUMPAD4"
    aKeys(&H4C) = "DIK_NUMPAD5"
    aKeys(&H4D) = "DIK_NUMPAD6"
    aKeys(&H4E) = "DIK_ADD" ' + on numeric keypad
    aKeys(&H4F) = "DIK_NUMPAD1"
    aKeys(&H50) = "DIK_NUMPAD2"
    aKeys(&H51) = "DIK_NUMPAD3"
    aKeys(&H52) = "DIK_NUMPAD0"
    aKeys(&H53) = "DIK_DECIMAL" ' . on numeric keypad
    aKeys(&H56) = "DIK_OEM_102 < > | on UK/Germany keyboards"
    aKeys(&H57) = "DIK_F11"
    aKeys(&H58) = "DIK_F12"
    aKeys(&H64) = "DIK_F13 on (NEC PC98) "
    aKeys(&H65) = "DIK_F14 on (NEC PC98) "
    aKeys(&H66) = "DIK_F15 on (NEC PC98) "
    aKeys(&H70) = "DIK_KANA on (Japanese keyboard)"
    aKeys(&H73) = "DIK_ABNT_C1 / ? on Portugese (Brazilian) keyboards "
    aKeys(&H79) = "DIK_CONVERT on (Japanese keyboard)"
    aKeys(&H7B) = "DIK_NOCONVERT on (Japanese keyboard)"
    aKeys(&H7D) = "DIK_YEN on (Japanese keyboard)"
    aKeys(&H7E) = "DIK_ABNT_C2 on Numpad . on Portugese (Brazilian) keyboards "
    aKeys(&H8D) = "DIK_NUMPADEQUALS = on numeric keypad (NEC PC98) "
    aKeys(&H90) = "DIK_PREVTRACK on Previous Track (DIK_CIRCUMFLEX on Japanese keyboard) "
    aKeys(&H91) = "DIK_AT (NEC PC98) "
    aKeys(&H92) = "DIK_COLON (NEC PC98) "
    aKeys(&H93) = "DIK_UNDERLINE (NEC PC98) "
    aKeys(&H94) = "DIK_KANJI on (Japanese keyboard)"
    aKeys(&H95) = "DIK_STOP (NEC PC98) "
    aKeys(&H96) = "DIK_AX (Japan AX) "
    aKeys(&H97) = "DIK_UNLABELED (J3100) "
    aKeys(&H99) = "DIK_NEXTTRACK" ' Next Track
    aKeys(&H9C) = "DIK_NUMPADENTER" ' Enter on numeric keypad
    aKeys(&H9D) = "DIK_RCONTROL"
    aKeys(&HA0) = "DIK_MUTE" ' Mute
    aKeys(&HA1) = "DIK_CALCULATOR" ' Calculator
    aKeys(&HA2) = "DIK_PLAYPAUSE" ' Play / Pause
    aKeys(&HA4) = "DIK_MEDIASTOP" ' Media Stop
    aKeys(&HAE) = "DIK_VOLUMEDOWN" ' Volume -
    aKeys(&HB0) = "DIK_VOLUMEUP" ' Volume +
    aKeys(&HB2) = "DIK_WEBHOME" ' Web home
    aKeys(&HB3) = "DIK_NUMPADCOMMA" ' , on numeric keypad (NEC PC98)
    aKeys(&HB5) = "DIK_DIVIDE" ' / on numeric keypad
    aKeys(&HB7) = "DIK_SYSRQ"
    aKeys(&HB8) = "DIK_RMENU" ' right Alt
    aKeys(&HC5) = "DIK_PAUSE" ' Pause
    aKeys(&HC7) = "DIK_HOME" ' Home on arrow keypad
    aKeys(&HC8) = "DIK_UP" ' UpArrow on arrow keypad
    aKeys(&HC9) = "DIK_PRIOR" ' PgUp on arrow keypad
    aKeys(&HCB) = "DIK_LEFT" ' LeftArrow on arrow keypad
    aKeys(&HCD) = "DIK_RIGHT" ' RightArrow on arrow keypad
    aKeys(&HCF) = "DIK_END" ' End on arrow keypad
    aKeys(&HD0) = "DIK_DOWN" ' DownArrow on arrow keypad
    aKeys(&HD1) = "DIK_NEXT" ' PgDn on arrow keypad
    aKeys(&HD2) = "DIK_INSERT" ' Insert on arrow keypad
    aKeys(&HD3) = "DIK_DELETE" ' Delete on arrow keypad
    aKeys(&HDB) = "DIK_LWIN" ' Left Windows key
    aKeys(&HDC) = "DIK_RWIN" ' Right Windows key
    aKeys(&HDD) = "DIK_APPS" ' AppMenu key
    aKeys(&HDE) = "DIK_POWER" ' System Power
    aKeys(&HDF) = "DIK_SLEEP" ' System Sleep
    aKeys(&HE3) = "DIK_WAKE" ' System Wake
    aKeys(&HE5) = "DIK_WEBSEARCH" ' Web Search
    aKeys(&HE6) = "DIK_WEBFAVORITES" ' Web Favorites
    aKeys(&HE7) = "DIK_WEBREFRESH" ' Web Refresh
    aKeys(&HE8) = "DIK_WEBSTOP" ' Web Stop
    aKeys(&HE9) = "DIK_WEBFORWARD" ' Web Forward
    aKeys(&HEA) = "DIK_WEBBACK" ' Web Back
    aKeys(&HEB) = "DIK_MYCOMPUTER" ' My Computer
    aKeys(&HEC) = "DIK_MAIL" ' Mail
    aKeys(&HED) = "DIK_MEDIASELECT" '

End Sub

Function ReCreateInputDevices() As Boolean

  ' Create Direct Input

  Dim K As Integer
  Dim I As Integer

    If Data.IS_DinputOK Then
        'free input device
        Set obj_Dinput = Nothing

        'On Error Resume Next
        DIKeyBoardDevice.Unacquire
        'log_out2 "              DirectInput Keyboard Device released.................OK"

        DIMouseDevice.Unacquire
        'log_out2 "              DirectInput Mouse Device released.................OK"

        If Data.IS_Joystick Then

            For I = 0 To Data.JoyNumDevice - 1
                DIjoyDevice(I).Unacquire
                Set DIjoyDevice(I) = Nothing
                'log_out2 "              DirectInput JoyPad Device N°" + STR(I) + " released.................OK"

            Next I
            Erase DIjoyDevice
            Erase AxisPresent

        End If

        Set DIKeyBoardDevice = Nothing
        Set DIMouseDevice = Nothing
    End If

    'log_out2 "<br>"
    'log_out2 "  ................INPUT DEVICES CREATION.............      '"
    'log_out2 "<br>"

    Set obj_Dinput = obj_DX.DirectInputCreate()

    'log_out2 "Main DirectInput Interface Created.......OK"

    ' Create keyboard device
    'log_out2 ""
    'log_out2 "............Keyboard.........."
    'log_out2 ""
    Set DIKeyBoardDevice = obj_Dinput.CreateDevice("GUID_SysKeyboard")
    ' Set common data format to keyboard

    'log_out2 "DirectInput Keyboard Device Created.......OK"

    Call DIKeyBoardDevice.SetCommonDataFormat(DIFORMAT_KEYBOARD)
    

    If CFG.IS_FullScreen Then
        Call DIKeyBoardDevice.SetCooperativeLevel(Data.HwindowParent, DISCL_EXCLUSIVE Or DISCL_FOREGROUND)

      Else
        Call DIKeyBoardDevice.SetCooperativeLevel(Data.HwindowParent, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)

    End If

    Call DIKeyBoardDevice.Acquire

    For K = 1 To 10
        'While Err.Number = DIERR_INPUTLOST
        Call DIKeyBoardDevice.Acquire

        'Wend
    Next K

    'log_out2 "DirectInput Keyboard Device Acquired.......OK"

    'log_out2 ""
    'log_out2 "............Mouse.........."
    'log_out2 ""
    ' Create Mouse device
    Set DIMouseDevice = obj_Dinput.CreateDevice("GUID_SysMouse")

    'log_out2 "DirectInput Mouse Device Created.......OK"

    'Set common data format to mouse
    DIMouseDevice.SetCommonDataFormat DIFORMAT_MOUSE
    'DIMouseDevice.SetCooperativeLevel Data.HwindowParent, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

    If CFG.IS_FullScreen Then
        DIMouseDevice.SetCooperativeLevel Data.HwindowParent, DISCL_EXCLUSIVE Or DISCL_FOREGROUND

      Else
        DIMouseDevice.SetCooperativeLevel Data.HwindowParent, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    End If

    DIMouseDevice.Acquire

    'log_out2 "DirectInput Mouse Device Acquired.......OK"

    Data.IS_DinputOK = True

    'create Joystick Device

    'log_out2 ""
    'log_out2 "............JoyPad.........."
    'log_out2 ""

    Set diDevEnumJoy = obj_Dinput.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)
    Set diDevEnumMouse = obj_Dinput.GetDIDevices(DI8DEVTYPE_MOUSE, DIEDFL_ATTACHEDONLY)
    Set diDevEnumKey = obj_Dinput.GetDIDevices(DI8DEVCLASS_KEYBOARD, DIEDFL_ATTACHEDONLY)

    Set diDevEnumAll = obj_Dinput.GetDIDevices(DI8DEVCLASS_ALL, DIEDFL_ATTACHEDONLY)

    If diDevEnumJoy.GetCount = 0 Then
        'log_out2 "No joystick attached.", QUEST3D_LOGWARNING
        Data.IS_Joystick = False
        Exit Function
      Else
        'log_out2 "........." + STR(diDevEnumJoy.GetCount) + " attached joystick found .."
        Data.JoyNumDevice = diDevEnumJoy.GetCount
        Data.IS_Joystick = True

    End If

    ReDim AxisPresent(K, Data.JoyNumDevice - 1, 1 To 8)

    ReDim DIjoyDevice(Data.JoyNumDevice - 1)
    ReDim joyCaps(Data.JoyNumDevice - 1)

    For K = 0 To Data.JoyNumDevice - 1
        Set DIjoyDevice(K) = obj_Dinput.CreateDevice(diDevEnumJoy.GetItem(K + 1).GetGuidInstance)

        'log_out2 "DirectInput JoyPad Device N°" + STR(K) + " Created.......OK"

        DIjoyDevice(K).SetCommonDataFormat DIFORMAT_JOYSTICK
        DIjoyDevice(K).SetCooperativeLevel Data.HwindowParent, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

        DIjoyDevice(K).GetCapabilities joyCaps(K)

        ' Set deadzone for X and Y axis to 10 percent of the range of travel
        With Data.DiProp_Dead
            .lData = 1000
            .lHow = DIPH_BYOFFSET

            .lObj = DIJOFS_X
            DIjoyDevice(K).SetProperty "DIPROP_DEADZONE", Data.DiProp_Dead

            .lObj = DIJOFS_Y
            DIjoyDevice(K).SetProperty "DIPROP_DEADZONE", Data.DiProp_Dead

        End With

        ' Set saturation zones for X and Y axis to 5 percent of the range
        With Data.DiProp_Saturation
            .lData = 9500
            .lHow = DIPH_BYOFFSET

            .lObj = DIJOFS_X
            DIjoyDevice(K).SetProperty "DIPROP_SATURATION", Data.DiProp_Saturation

            .lObj = DIJOFS_Y
            DIjoyDevice(K).SetProperty "DIPROP_SATURATION", Data.DiProp_Saturation

        End With

        ' NOTE Some devices do not let you set the range

        ' Set range for all axes
        With Data.DiProp_Range
            .lHow = DIPH_DEVICE
            .lMin = -1000
            .lMax = 1000
        End With

        On Error Resume Next
            DIjoyDevice(K).SetProperty "DIPROP_RANGE", Data.DiProp_Range

            DIjoyDevice(K).Acquire

            If Not (DIjoyDevice(K) Is Nothing) Then Data.IS_Joystick = True

      Dim didoEnum As DirectInputEnumDeviceObjects
      Dim dido As DirectInputDeviceObjectInstance
     

            For I = 1 To 8
                AxisPresent(K, I) = False
            Next I

            ' Enumerate the axes
            Set didoEnum = DIjoyDevice(K).GetDeviceObjectsEnum(DIDFT_AXIS)

            ' Check data offset of each axis to learn what it is
      Dim sGuid As String
            For I = 1 To didoEnum.GetCount

                Set dido = didoEnum.GetItem(I)

                sGuid = dido.GetGuidType
                Select Case sGuid
                  Case "GUID_XAxis"
                    AxisPresent(K, 1) = True

                  Case "GUID_YAxis"
                    AxisPresent(K, 2) = True

                  Case "GUID_ZAxis"
                    AxisPresent(K, 3) = True
                    ''log_out2 "Z_AXIS found for JoyPad Device " + STR(K + 1)
                  Case "GUID_RxAxis"
                    AxisPresent(K, 4) = True
                    ''log_out2 "Rx_AXIS found for JoyPad Device " + STR(K + 1)

                  Case "GUID_RyAxis"
                    AxisPresent(K, 5) = True
                    'log_out2 "Ry_AXIS found for JoyPad Device " + STR(K + 1)

                  Case "GUID_RzAxis"
                    AxisPresent(K, 6) = True
                    ''log_out2 "Rz_AXIS found for JoyPad Device " + STR(K + 1)

                  Case "GUID_Slider"
                    AxisPresent(K, 8) = True
                    AxisPresent(K, 7) = True
                    ''log_out2 "Slider found for JoyPad Device " + STR(K + 1)

                End Select

            Next I

        Next K

        'log_out2 "<br>"

End Function

Function Get_DirectInput(Optional GamePadID As Long = 1) As DirectInput8

    Set Get_DirectInput = obj_Dinput

End Function

Function Get_DirectInputEnumDevices(Optional DevType As QUEST3D_INPUTDEVICEICETYPE) As DirectInputEnumDevices8

    If DevType = QUEST3D_INPUTDEVICE_CLASS_ALL Then
        Set Get_DirectInputEnumDevices = diDevEnumAll
        Exit Function
      ElseIf DevType = QUEST3D_INPUTDEVICE_CLASS_GAMECTRL Then
        Set Get_DirectInputEnumDevices = diDevEnumJoy
        Exit Function
      ElseIf DevType = QUEST3D_INPUTDEVICE_CLASS_KEYBOARD Then
        Set Get_DirectInputEnumDevices = diDevEnumKey
        Exit Function

      ElseIf DevType = QUEST3D_INPUTDEVICE_TYPE_MOUSE Then
        Set Get_DirectInputEnumDevices = diDevEnumMouse
        Exit Function
    End If

End Function

Function Get_GamepadCapabilities(Optional GamePadID As Long = 1) As DIDEVCAPS

    If Not Data.IS_Joystick Then Exit Function
    Get_GamepadCapabilities = joyCaps(GamePadID - 1)

End Function

Function Get_GamePadInputDevice(Optional GamePadID As Long = 1) As DirectInputDevice8

    Set Get_GamePadInputDevice = DIjoyDevice(GamePadID)

End Function

Function Get_GamePadState(ByVal GamePadButton As QUEST3D_GAMEPAD_STATE, Optional ByVal GamePadID As Long = 1) As QUEST3D_KEY_STATE

  Dim JS As DIJOYSTATE

    On Local Error Resume Next

        DIjoyDevice(GamePadID - 1).Poll
        DIjoyDevice(GamePadID - 1).GetDeviceStateJoystick JS
        If Err.Number = DIERR_NOTACQUIRED Or Err.Number = DIERR_INPUTLOST Then
            DIjoyDevice(GamePadID - 1).Acquire
            Exit Function
        End If

        If GamePadButton >= 48 Then
            If (JS.Buttons(GamePadButton - 48) > 0) Then Get_GamePadState = QUEST3D_KEYDOWN
            Exit Function
        End If

        If GamePadButton = QUEST3D_GAMEPAD_RIGHT Then
            Get_GamePadState = (JS.x >= Data.DiProp_Range.lMax / 2) And Not (JS.y <= Data.DiProp_Range.lMin / 2) And Not (JS.y >= Data.DiProp_Range.lMax / 2)
            Exit Function
          ElseIf GamePadButton = QUEST3D_GAMEPAD_DIAG_UPRIGHT Then
            Get_GamePadState = (JS.x > Data.DiProp_Range.lMax / 2) And (JS.y <= Data.DiProp_Range.lMin / 2)
            Exit Function

          ElseIf GamePadButton = QUEST3D_GAMEPAD_DIAG_DOWNRIGHT Then
            Get_GamePadState = (JS.x > Data.DiProp_Range.lMax / 2) And (JS.y >= Data.DiProp_Range.lMax / 2)
            Exit Function

          ElseIf GamePadButton = QUEST3D_GAMEPAD_DOWN Then
            Get_GamePadState = (JS.y >= Data.DiProp_Range.lMax / 2) And Not (JS.x >= Data.DiProp_Range.lMax / 2) And Not (JS.x <= Data.DiProp_Range.lMin / 2)
            Exit Function

          ElseIf GamePadButton = QUEST3D_GAMEPAD_DIAG_DOWNLEFT Then
            Get_GamePadState = (JS.x <= Data.DiProp_Range.lMin / 2) And (JS.y >= Data.DiProp_Range.lMax / 2)
            Exit Function

          ElseIf GamePadButton = QUEST3D_GAMEPAD_LEFT Then
            Get_GamePadState = (JS.x <= Data.DiProp_Range.lMin / 2) And Not (JS.y <= Data.DiProp_Range.lMin / 2) And Not (JS.y >= Data.DiProp_Range.lMax / 2)
            Exit Function

          ElseIf GamePadButton = QUEST3D_GAMEPAD_UP Then
            Get_GamePadState = (JS.y <= Data.DiProp_Range.lMin / 2) And Not (JS.x >= Data.DiProp_Range.lMax / 2) And Not (JS.x <= Data.DiProp_Range.lMin / 2)
            Exit Function

          ElseIf GamePadButton = QUEST3D_GAMEPAD_DIAG_UPLEFT Then
            Get_GamePadState = (JS.x <= Data.DiProp_Range.lMin / 2) And (JS.y <= Data.DiProp_Range.lMin / 2)
            Exit Function

        End If

End Function

Function Get_GamePadState2(Optional GamePadID As Long = 1) As DIJOYSTATE

  Dim JS As DIJOYSTATE

    On Local Error Resume Next

        DIjoyDevice(GamePadID - 1).Poll
        DIjoyDevice(GamePadID - 1).GetDeviceStateJoystick JS
        If Err.Number = DIERR_NOTACQUIRED Or Err.Number = DIERR_INPUTLOST Then
            DIjoyDevice(GamePadID - 1).Acquire
            Exit Function
        End If

        Get_GamePadState2 = JS

End Function

Function Get_GameProperty(ByVal GuidProperty As String, Optional ByVal GamePadID As Long = 1) As DIPROPLONG

    DIjoyDevice(GamePadID - 1).GetProperty GuidProperty, Get_GameProperty

End Function

'============================================================
'
'KEYBOARD and MOUSE DEVICE
'
'===========================================================

'----------------------------------------
'Name: Get_KeyPress
'Object: Get
'Event: KeyPress
'----------------------------------------
'----------------------------------------
'Name: Get_KeyPress
'Object: Get
'Event: KeyPress
'Description:
'----------------------------------------
Function Get_KeyBoardKeyPressed(KEY As QUEST3D_KEY_CONST) As Boolean

  '    If Data.IS_DinputOK = False Then
  '        Dim P As Integer
  '        P = GetAsyncKeyState(KEY)
  '        Get_KeyBoardKeyPressed = P
  '        Exit Function
  '    End If

    If Data.IS_DinputOK = False Then Exit Function

    DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE

    '    'If Err.Number = DIERR_INPUTLOST Then
    '            DIKeyBoardDevice.Acquire
    '            DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE
    '            DIKeyBoardDevice.Poll
    '   ' End If

    ' If key is pressed then
    If DIKEYBOARDSTATE.KEY(KEY) <> 0 Then Get_KeyBoardKeyPressed = True

End Function

Sub ReAcquireInputDevices(Optional ByVal ForceInputDevice As Boolean = True)

  Dim I As Integer
  Dim J As Integer
  Dim K As Integer

    If ForceInputDevice Then
        K = 10
      Else
        K = 1
    End If

    For J = 1 To K
        If Data.IS_DinputOK Then

            DIKeyBoardDevice.Poll
            DIKeyBoardDevice.Acquire
            DIMouseDevice.Acquire
        End If

        If Data.IS_Joystick Then

            For I = 0 To Data.JoyNumDevice - 1
                DIjoyDevice(I).Acquire
            Next I
        End If
    Next J

End Sub

Function Get_KeyBoardKeyPressedAPI(KEY As KeyCodeConstants) As Boolean

    Get_KeyBoardKeyPressedAPI = GetAsyncKeyState(KEY)

End Function

Function Get_KeyBoardKeysPressedNames() As String

  Dim s As String
  Dim Tcount As Integer

    DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE

    Tcount = 0

    While Tcount < 256
        ' If key is pressed then
        If DIKEYBOARDSTATE.KEY(Tcount) <> 0 Then
            s = s + aKeys(Tcount) + Chr(10)
        End If
        Tcount = Tcount + 1
    Wend

    Get_KeyBoardKeysPressedNames = s

End Function

Function Get_MOUSE_Cursor_API_X() As Integer

  Dim Pos As POINTAPI

    Call GetCursorPos(Pos)

    If Data.Is_Windowed Then
        'Dim rc As RECT
        'Call GetWindowRect(Data.Hwindow, rc)

        Pos.x = Pos.x - Data.Buffer_Rect.Left

    End If

    Get_MOUSE_Cursor_API_X = Pos.x

End Function

Function Get_MOUSE_Cursor_API_Y() As Integer

  Dim Pos As POINTAPI

    Call GetCursorPos(Pos)

    If Data.Is_Windowed Then
        'Dim rc As RECT
        'Call GetWindowRect(Data.Hwindow, rc)

        Pos.y = Pos.y - Data.Buffer_Rect.Top
    End If
    Get_MOUSE_Cursor_API_Y = Pos.y

End Function

'

'End Function

'----------------------------------------
'Name: Get_MouseBouton
'Object: Get
'Event: MouseBouton
'----------------------------------------
'----------------------------------------
'Name: Get_MouseBouton
'Object: Get
'Event: MouseBouton
'Description:
'----------------------------------------
Function Get_MouseBouton(BUTTON As QUEST3D_MOUSE_BUTTON) As Boolean

  ' Lets get the mouse state

    DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE
    If DIMOUSESTATE.Buttons(BUTTON) <> 0 Then Get_MouseBouton = True

End Function

Function Get_MouseWheel() As Long

    DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE
    Get_MouseWheel = DIMOUSESTATE.lZ

End Function

Function Get_MouseDevice() As DirectInputDevice8

    Set Get_MouseDevice = DIMouseDevice

End Function

Function Get_MouseRect() As RECT
    
    Dim Pos As POINTAPI

    Call GetCursorPos(Pos)

    If Data.Is_Windowed Then
        'Dim rc As RECT
        'Call GetWindowRect(Data.Hwindow, rc)

        Pos.x = Pos.x - Data.Buffer_Rect.Left
        Pos.y = Pos.y - Data.Buffer_Rect.Top

    End If
  Get_MouseRect.Left = Pos.x
  Get_MouseRect.Top = Pos.y
  Get_MouseRect.Right = Pos.x + 8
  Get_MouseRect.bottom = Pos.y + 15
  
  
End Function

Function Get_MousePos() As D3DVECTOR2

  Dim Pos As POINTAPI

    Call GetCursorPos(Pos)

    If Data.Is_Windowed Then
        'Dim rc As RECT
        'Call GetWindowRect(Data.Hwindow, rc)

        Pos.x = Pos.x - Data.Buffer_Rect.Left
        Pos.y = Pos.y - Data.Buffer_Rect.Top

    End If

    Get_MousePos.x = Pos.x
    Get_MousePos.y = Pos.y

End Function



'----------------------------------------
'Name: Get_MouseXmove
'Object: Get
'Event: MouseXmove
'----------------------------------------
'----------------------------------------
'Name: Get_MouseXmove
'Object: Get
'Event: MouseXmove
'Description:
'----------------------------------------
Function Get_MouseXmove() As Integer

  ' Lets get the mouse state

    DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE
    Get_MouseXmove = DIMOUSESTATE.lX

End Function

Function Get_MouseMove() As D3DVECTOR2

  ' Lets get the mouse state

    DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE
    Get_MouseMove.x = DIMOUSESTATE.lX
    Get_MouseMove.y = DIMOUSESTATE.lY

End Function


'----------------------------------------
'Name: Get_MouseYmove
'Object: Get
'Event: MouseYmove
'----------------------------------------
'----------------------------------------
'Name: Get_MouseYmove
'Object: Get
'Event: MouseYmove
'Description:
'----------------------------------------
Function Get_MouseYmove() As Integer

  ' Lets get the mouse state

    DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE
    Get_MouseYmove = DIMOUSESTATE.lY

End Function

Function Get_NumGamePadAttached() As Long

    Get_NumGamePadAttached = Data.JoyNumDevice

End Function

Sub Set_GameDeadProperty(ByVal value As Long, ByVal DeviceOBJ As CONST_DIJOYSTICKOFS, Optional ByVal lHow As CONST_DIPHFLAGS = DIPH_BYOFFSET, Optional ByVal GamePadID As Long = 1)

  ' Set range for all axes
  'Dead As DIPROPLONG

  ' Set deadzone for X and Y axis to 10 percent of the range of travel

    With Data.DiProp_Dead
        .lData = value
        .lHow = lHow

        .lObj = DeviceOBJ
        DIjoyDevice(GamePadID - 1).SetProperty "DIPROP_DEADZONE", Data.DiProp_Dead

    End With

End Sub

Sub Set_GameDeadPropertyEX(DiProp As DIPROPLONG, Optional ByVal GamePadID As Long = 1)

  ' Set range for all axes
  'Dead As DIPROPLONG

  ' Set deadzone for X and Y axis to 10 percent of the range of travel

    With Data.DiProp_Dead

        DIjoyDevice(GamePadID - 1).SetProperty "DIPROP_DEADZONE", Data.DiProp_Dead

    End With

End Sub

Sub Set_GamePadRangeProperty(lMax As Long, lMin As Long, Optional ByVal lHow As CONST_DIPHFLAGS = DIPH_DEVICE, Optional ByVal GamePadID As Long = 1)

  ' Set range for all axes
  'Range As DIPROPRANGE

    With Data.DiProp_Range
        .lHow = lHow
        .lMin = lMin
        .lMax = lMax

    End With
    DIjoyDevice(GamePadID - 1).SetProperty "DIPROP_RANGE", Data.DiProp_Range

End Sub

Sub Set_GamePadRangePropertyEX(DinputProp_Range As DIPROPRANGE, Optional ByVal GamePadID As Long = 1)

  ' Set range for all axes

    DIjoyDevice(GamePadID - 1).SetProperty "DIPROP_RANGE", DinputProp_Range

End Sub

Sub Set_GameSaturationProperty(ByVal value As Long, DeviceOBJ As CONST_DIJOYSTICKOFS, Optional ByVal lHow As CONST_DIPHFLAGS = DIPH_BYOFFSET, Optional ByVal GamePadID As Long = 1)

  'Saturation As DIPROPLONG

  ' Set deadzone for X and Y axis to 10 percent of the range of travel

    With Data.DiProp_Dead
        .lData = value
        .lHow = lHow

        .lObj = DeviceOBJ
        DIjoyDevice(GamePadID - 1).SetProperty "DIPROP_SATURATION", Data.DiProp_Dead

    End With

End Sub

Sub Set_GameSaturationPropertyEX(DinputProp_Saturation As DIPROPLONG, Optional ByVal GamePadID As Long = 1)

    DIjoyDevice(GamePadID - 1).SetProperty "DIPROP_SATURATION", DinputProp_Saturation

End Sub

