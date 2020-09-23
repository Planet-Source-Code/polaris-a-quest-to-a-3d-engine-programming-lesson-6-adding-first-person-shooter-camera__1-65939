Attribute VB_Name = "Module_Input"
'========Input========
'keyboard
Public obj_Dinput As DirectInput8 'this is DirectInput, used to monitor the keys on the keyboard in my case
Public DIKeyBoardDevice As DirectInputDevice8 'this device will be the keyboard
Public DIKEYBOARDSTATE As DIKEYBOARDSTATE 'to check the state of keys

Public DIMouseDevice As DirectInputDevice8 ' Mouse device
Public DIMOUSESTATE As DIMOUSESTATE ' to check mouse movements and clicks

Public DIjoyDevice() As DirectInputDevice8
Public diDevEnumJoy As DirectInputEnumDevices8
Public diDevEnumMouse As DirectInputEnumDevices8
Public diDevEnumKey As DirectInputEnumDevices8
Public diDevEnumAll As DirectInputEnumDevices8

Public EventHandle As Long
Public joyCaps() As DIDEVCAPS
Public JS As DIJOYSTATE
'Public DiProp_Dead As DIPROPLONG
'Public DiProp_Range As DIPROPRANGE
'Public DiProp_Saturation As DIPROPLONG
Public AxisPresent() As Boolean


'this constants are from DirectX 8 SDK
Public Enum QUEST3D_KEY_CONST

        QUEST3D_KEY_ESCAPE = &H1            '
        QUEST3D_KEY_1 = &H2                 '
        QUEST3D_KEY_2 = &H3                 '
        QUEST3D_KEY_3 = &H4                 '
        QUEST3D_KEY_4 = &H5                 '
        QUEST3D_KEY_5 = &H6                 '
        QUEST3D_KEY_6 = &H7                 '
        QUEST3D_KEY_7 = &H8                 '
        QUEST3D_KEY_8 = &H9                 '
        QUEST3D_KEY_9 = &HA                 '
        QUEST3D_KEY_0 = &HB                 '
        QUEST3D_KEY_MINUS = &HC             '    /* - on main keyboard */
        QUEST3D_KEY_EQUALS = &HD            '
        QUEST3D_KEY_BACK = &HE              '    /* backspace */
        QUEST3D_KEY_TAB = &HF               '
        QUEST3D_KEY_Q = &H10                '
        QUEST3D_KEY_W = &H11                '
        QUEST3D_KEY_E = &H12                '
        QUEST3D_KEY_R = &H13                '
        QUEST3D_KEY_T = &H14                '
        QUEST3D_KEY_Y = &H15                '
        QUEST3D_KEY_U = &H16                '
        QUEST3D_KEY_I = &H17                '
        QUEST3D_KEY_O = &H18                '
        QUEST3D_KEY_P = &H19                '
        QUEST3D_KEY_LBRACKET = &H1A         '
        QUEST3D_KEY_RBRACKET = &H1B         '
        QUEST3D_KEY_RETURN = &H1C           '    /* Enter on main keyboard */
        QUEST3D_KEY_LCONTROL = &H1D         '
        QUEST3D_KEY_A = &H1E                '
        QUEST3D_KEY_S = &H1F                '
        QUEST3D_KEY_D = &H20                '
        QUEST3D_KEY_F = &H21                '
        QUEST3D_KEY_G = &H22                '
        QUEST3D_KEY_H = &H23                '
        QUEST3D_KEY_J = &H24                '
        QUEST3D_KEY_K = &H25                '
        QUEST3D_KEY_L = &H26                '
        QUEST3D_KEY_SEMICOLON = &H27        '
        QUEST3D_KEY_APOSTROPHE = &H28       '
        QUEST3D_KEY_GRAVE = &H29            '    /* accent grave */
        QUEST3D_KEY_LSHIFT = &H2A           '
        QUEST3D_KEY_BACKSLASH = &H2B        '
        QUEST3D_KEY_Z = &H2C                '
        QUEST3D_KEY_X = &H2D                '
        QUEST3D_KEY_C = &H2E                '
        QUEST3D_KEY_V = &H2F                '
        QUEST3D_KEY_B = &H30                '
        QUEST3D_KEY_N = &H31                '
        QUEST3D_KEY_M = &H32                '
        QUEST3D_KEY_COMMA = &H33            '
        QUEST3D_KEY_PERIOD = &H34           '    /* . on main keyboard */
        QUEST3D_KEY_SLASH = &H35            '    /* / on main keyboard */
        QUEST3D_KEY_RSHIFT = &H36           '
        QUEST3D_KEY_MULTIPLY = &H37         '    /* * on numeric keypad */
        QUEST3D_KEY_LMENU = &H38            '    /* left Alt */
        QUEST3D_KEY_SPACE = &H39            '
        QUEST3D_KEY_CAPITAL = &H3A          '
        QUEST3D_KEY_F1 = &H3B               '
        QUEST3D_KEY_F2 = &H3C               '
        QUEST3D_KEY_F3 = &H3D               '
        QUEST3D_KEY_F4 = &H3E               '
        QUEST3D_KEY_F5 = &H3F               '
        QUEST3D_KEY_F6 = &H40               '
        QUEST3D_KEY_F7 = &H41               '
        QUEST3D_KEY_F8 = &H42               '
        QUEST3D_KEY_F9 = &H43               '
        QUEST3D_KEY_F10 = &H44              '
        QUEST3D_KEY_NUMLOCK = &H45          '
        QUEST3D_KEY_SCROLL = &H46           '    /* Scroll Lock */
        QUEST3D_KEY_NUMPAD7 = &H47          '
        QUEST3D_KEY_NUMPAD8 = &H48          '
        QUEST3D_KEY_NUMPAD9 = &H49          '
        QUEST3D_KEY_SUBTRACT = &H4A         '    /* - on numeric keypad */
        QUEST3D_KEY_NUMPAD4 = &H4B          '
        QUEST3D_KEY_NUMPAD5 = &H4C          '
        QUEST3D_KEY_NUMPAD6 = &H4D          '
        QUEST3D_KEY_ADD = &H4E              '    /* + on numeric keypad */
        QUEST3D_KEY_NUMPAD1 = &H4F          '
        QUEST3D_KEY_NUMPAD2 = &H50          '
        QUEST3D_KEY_NUMPAD3 = &H51          '
        QUEST3D_KEY_NUMPAD0 = &H52          '
        QUEST3D_KEY_DECIMAL = &H53          '    /* . on numeric keypad */
        QUEST3D_KEY_OEM_102 = &H56          '    /* <> or \| on RT 102-key keyboard (Non-U.S.) */
        QUEST3D_KEY_F11 = &H57              '
        QUEST3D_KEY_F12 = &H58              '
        QUEST3D_KEY_F13 = &H64              '    /*                     (NEC PC98) */
        QUEST3D_KEY_F14 = &H65              '    /*                     (NEC PC98) */
        QUEST3D_KEY_F15 = &H66              '    /*                     (NEC PC98) */
        QUEST3D_KEY_KANA = &H70             '    /* (Japanese keyboard)            */
        QUEST3D_KEY_ABNT_C1 = &H73          '    /* /? on Brazilian keyboard */
        QUEST3D_KEY_CONVERT = &H79          '    /* (Japanese keyboard)            */
        QUEST3D_KEY_NOCONVERT = &H7B        '    /* (Japanese keyboard)            */
        QUEST3D_KEY_YEN = &H7D              '    /* (Japanese keyboard)            */
        QUEST3D_KEY_ABNT_C2 = &H7E          '    /* Numpad . on Brazilian keyboard */
        QUEST3D_KEY_NUMPADEQUALS = &H8D     '    /* = on numeric keypad (NEC PC98) */
        QUEST3D_KEY_PREVTRACK = &H90        '    /* Previous Track (DIK_CIRCUMFLEX on Japanese keyboard) */
        QUEST3D_KEY_AT = &H91               '    /*                     (NEC PC98) */
        QUEST3D_KEY_COLON = &H92            '    /*                     (NEC PC98) */
        QUEST3D_KEY_UNDERLINE = &H93        '    /*                     (NEC PC98) */
        QUEST3D_KEY_KANJI = &H94            '    /* (Japanese keyboard)            */
        QUEST3D_KEY_STOP = &H95             '    /*                     (NEC PC98) */
        QUEST3D_KEY_AX = &H96               '    /*                     (Japan AX) */
        QUEST3D_KEY_UNLABELED = &H97        '    /*                        (J3100) */
        QUEST3D_KEY_NEXTTRACK = &H99        '    /* Next Track */
        QUEST3D_KEY_NUMPADENTER = &H9C      '    /* Enter on numeric keypad */
        QUEST3D_KEY_RCONTROL = &H9D         '
        QUEST3D_KEY_MUTE = &HA0             '    /* Mute */
        QUEST3D_KEY_CALCULATOR = &HA1       '    /* Calculator */
        QUEST3D_KEY_PLAYPAUSE = &HA2        '    /* Play / Pause */
        QUEST3D_KEY_MEDIASTOP = &HA4        '    /* Media Stop */
        QUEST3D_KEY_VOLUMEDOWN = &HAE       '    /* Volume - */
        QUEST3D_KEY_VOLUMEUP = &HB0         '    /* Volume + */
        QUEST3D_KEY_WEBHOME = &HB2          '    /* Web home */
        QUEST3D_KEY_NUMPADCOMMA = &HB3      '    /*    ' on numeric keypad (NEC PC98) */
        QUEST3D_KEY_DIVIDE = &HB5           '    /* / on numeric keypad */
        QUEST3D_KEY_SYSRQ = &HB7            '
        QUEST3D_KEY_RMENU = &HB8            '    /* right Alt */
        QUEST3D_KEY_PAUSE = &HC5            '    /* Pause */
        QUEST3D_KEY_HOME = &HC7             '    /* Home on arrow keypad */
        QUEST3D_KEY_UP = &HC8               '    /* UpArrow on arrow keypad */
        QUEST3D_KEY_PRIOR = &HC9            '    /* PgUp on arrow keypad */
        QUEST3D_KEY_LEFT = &HCB             '    /* LeftArrow on arrow keypad */
        QUEST3D_KEY_RIGHT = &HCD            '    /* RightArrow on arrow keypad */
        QUEST3D_KEY_END = &HCF              '    /* End on arrow keypad */
        QUEST3D_KEY_DOWN = &HD0             '    /* DownArrow on arrow keypad */
        QUEST3D_KEY_NEXT = &HD1             '    /* PgDn on arrow keypad */
        QUEST3D_KEY_INSERT = &HD2           '    /* Insert on arrow keypad */
        QUEST3D_KEY_DELETE = &HD3           '    /* Delete on arrow keypad */
        QUEST3D_KEY_LWIN = &HDB             '    /* Left Windows key */
        QUEST3D_KEY_RWIN = &HDC             '    /* Right Windows key */
        QUEST3D_KEY_APPS = &HDD             '    /* AppMenu key */
        QUEST3D_KEY_POWER = &HDE            '    /* System Power */
        QUEST3D_KEY_SLEEP = &HDF            '    /* System Sleep */
        QUEST3D_KEY_WAKE = &HE3             '    /* System Wake */
        QUEST3D_KEY_WEBSEARCH = &HE5        '    /* Web Search */
        QUEST3D_KEY_WEBFAVORITES = &HE6     '    /* Web Favorites */
        QUEST3D_KEY_WEBREFRESH = &HE7       '    /* Web Refresh */
        QUEST3D_KEY_WEBSTOP = &HE8          '    /* Web Stop */
        QUEST3D_KEY_WEBFORWARD = &HE9       '    /* Web Forward */
        QUEST3D_KEY_WEBBACK = &HEA          '    /* Web Back */
        QUEST3D_KEY_MYCOMPUTER = &HEB       '    /* My Computer */
        QUEST3D_KEY_MAIL = &HEC             '    /* Mail */
        QUEST3D_KEY_MEDIASELECT = &HED      '    /* Media Select */


End Enum










'----------------------------------------
'
'This sub Initialize All Input Device
'Mouse
'Keyboard
'Joystick
'----------------------------------------
Sub Init_Input(ByVal HAND As Long)

  ' Create Direct Input

  Dim K As Integer

    Set obj_Dinput = obj_DX.DirectInputCreate()

    ' Create keyboard device
   
    Set DIKeyBoardDevice = obj_Dinput.CreateDevice("GUID_SysKeyboard")
    ' Set common data format to keyboard
  
    
    

   

    Call DIKeyBoardDevice.SetCommonDataFormat(DIFORMAT_KEYBOARD)

    If CFG.IS_FullScreen Then
        Call DIKeyBoardDevice.SetCooperativeLevel(HAND, DISCL_EXCLUSIVE Or DISCL_FOREGROUND)

      Else
        Call DIKeyBoardDevice.SetCooperativeLevel(HAND, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)

    End If

    Call DIKeyBoardDevice.Acquire

For K = 1 To 10
    'While Err.Number = DIERR_INPUTLOST
        Call DIKeyBoardDevice.Acquire
  
     'Wend
Next K
  
   
   

   
    ' Create Mouse device
    Set DIMouseDevice = obj_Dinput.CreateDevice("GUID_SysMouse")

  
    'Set common data format to mouse
    DIMouseDevice.SetCommonDataFormat DIFORMAT_MOUSE
    'DIMouseDevice.SetCooperativeLevel HAND, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

    If CFG.IS_FullScreen Then
        DIMouseDevice.SetCooperativeLevel HAND, DISCL_EXCLUSIVE Or DISCL_FOREGROUND

      Else
        DIMouseDevice.SetCooperativeLevel HAND, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    End If

    DIMouseDevice.Acquire

  
    Data.IS_DinputOK = True

    'create Joystick Device

    Set diDevEnumJoy = obj_Dinput.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)
    Set diDevEnumMouse = obj_Dinput.GetDIDevices(DI8DEVTYPE_MOUSE, DIEDFL_ATTACHEDONLY)
    Set diDevEnumKey = obj_Dinput.GetDIDevices(DI8DEVCLASS_KEYBOARD, DIEDFL_ATTACHEDONLY)

    Set diDevEnumAll = obj_Dinput.GetDIDevices(DI8DEVCLASS_ALL, DIEDFL_ATTACHEDONLY)

    If diDevEnumJoy.GetCount = 0 Then
        Data.IS_Joystick = False
        Exit Sub
      Else
        Data.JoyNumDevice = diDevEnumJoy.GetCount
        Data.IS_Joystick = True

    End If

    ReDim AxisPresent(Data.JoyNumDevice - 1, 1 To 8)

    ReDim DIjoyDevice(Data.JoyNumDevice - 1)
    ReDim joyCaps(Data.JoyNumDevice - 1)

    For K = 0 To Data.JoyNumDevice - 1
        Set DIjoyDevice(K) = obj_Dinput.CreateDevice(diDevEnumJoy.GetItem(K + 1).GetGuidInstance)


        DIjoyDevice(K).SetCommonDataFormat DIFORMAT_JOYSTICK
        DIjoyDevice(K).SetCooperativeLevel HAND, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

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

        'On Error Resume Next
            DIjoyDevice(K).SetProperty "DIPROP_RANGE", Data.DiProp_Range

            DIjoyDevice(K).Acquire

            If Not (DIjoyDevice(K) Is Nothing) Then Data.IS_Joystick = True

      Dim didoEnum As DirectInputEnumDeviceObjects
      Dim dido As DirectInputDeviceObjectInstance
      Dim I As Integer

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
                    'log_out2 "Z_AXIS found for JoyPad Device " + STR(K + 1)
                  Case "GUID_RxAxis"
                    AxisPresent(K, 4) = True
                    'log_out2 "Rx_AXIS found for JoyPad Device " + STR(K + 1)

                  Case "GUID_RyAxis"
                    AxisPresent(K, 5) = True
                   
                  Case "GUID_RzAxis"
                    AxisPresent(K, 6) = True
                    'log_out2 "Rz_AXIS found for JoyPad Device " + STR(K + 1)

                  Case "GUID_Slider"
                    AxisPresent(K, 8) = True
                    AxisPresent(K, 7) = True
                    'log_out2 "Slider found for JoyPad Device " + STR(K + 1)

                End Select

            Next I

        Next K

        

       

End Sub


