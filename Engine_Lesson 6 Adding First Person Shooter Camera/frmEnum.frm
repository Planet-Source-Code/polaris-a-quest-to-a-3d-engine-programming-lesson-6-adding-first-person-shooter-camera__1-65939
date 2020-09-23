VERSION 5.00
Begin VB.Form frmEnum 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Engine Enumeration Dialogue"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00282D4A&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2205
      TabIndex        =   13
      Text            =   "1.00"
      Top             =   2010
      Width           =   2430
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00282D4A&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2205
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1290
      Width           =   3255
   End
   Begin VB.ComboBox NemoCmbAdapters 
      Appearance      =   0  'Flat
      BackColor       =   &H00282D4A&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2205
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   240
      Width           =   3255
   End
   Begin VB.ComboBox NemoCmbDevice 
      Appearance      =   0  'Flat
      BackColor       =   &H00282D4A&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2205
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   930
      Width           =   3255
   End
   Begin VB.ComboBox NemoCmbRes 
      Appearance      =   0  'Flat
      BackColor       =   &H00282D4A&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2205
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   585
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00282D4A&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2205
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1635
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "   START"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Vertical Sync Disabled"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Double Buffering"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   3000
      Width           =   2535
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tripple Buffering"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Rendering Mode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   2535
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Windowed"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "FullScreen"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         MaskColor       =   &H00008080&
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label GAMdown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4725
      TabIndex        =   22
      Top             =   2010
      Width           =   360
   End
   Begin VB.Label lblGammaUpDown 
      Alignment       =   2  'Center
      BackColor       =   &H00282D4A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5070
      TabIndex        =   21
      Top             =   930
      Width           =   390
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Gamma Level:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   495
      TabIndex        =   20
      Top             =   2010
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Displays available:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   270
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Resolution:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   615
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Acceleration: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Video Depth:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   16
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "BackBufferBit:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label GAMup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   14
      Top             =   2010
      Width           =   360
   End
End
Attribute VB_Name = "frmEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================================
'
'Here we enumerate all information about the GFX card
'  - screen resolutions supported like 640 x 480; 800 x 600; 1024 x 768
'  - device Type Hal=mean accelerated, or full hardware accelerated
'  -backbuffer color format like:
'       D3DFMT_R8G8B8
'        for "24-bit RGB pixel format."
'      D3DFMT_A8R8G8B8
'        for "32-bit ARGB pixel format with alpha."
'      D3DFMT_X8R8G8B8
'        for "32-bit" ' RGB"
'      D3DFMT_R5G6B5
'        for "16-bit" ' RGB"
'      D3DFMT_X1R5G5B5
'        for "16-bit pixel format where 5 bits are reserved for each color."
'      D3DFMT_A1R5G5B5
'        for "16-bit pixel format where 5 bits are reserved for each color and 1 bit is reserved for alpha (transparent texel)."
'      D3DFMT_A4R4G4B4
'        for "16-bit ARGB pixel format."
'      D3DFMT_R3G3B2
'        for "8-bit RGB texture format."
'
'
'
'=========================================================================================================

Option Explicit

Private CurrINDEX As Long

'Const VK_H = 72
'Const VK_E = 69
'Const VK_L = 76
'Const VK_O = 79
Const VK_RIGHT = &H27
'Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function SetFocu Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private GAMA_VAL As Single

Private Sub Combo2_Click()

    If DEV.CurrBKFMT = 0 Then Exit Sub
  Dim DV As tFMT

    If Option2.Value = True Then DV = DEV.CurrentWINDOWED Else DV = DEV.RESO(NemoCmbRes.ListIndex)

  Dim I, J, K
    K = Combo2.ListIndex
    'Combo2.Clear
    Combo1.Clear

    'If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
    If Option2.Value = 1 Then
        Combo1.Clear
        Call FillWindowedList

      Else
        Combo1.Clear
        For I = 0 To DV.DP_FMT(K).NumD - 1
            Combo1.AddItem FMTsTR(DV.DP_FMT(K).FMT(I))

        Next I
        If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
    End If

End Sub

Private Sub Command1_Click()

    Unload Me
    IS_ERROR = True

End Sub

Private Sub Command2_Click()

    Hide
    
    CFG.ForceVerSINC = Check1.Value
    If Option3.Value = 1 Then CFG.BufferCount = 2

    If Option4.Value = 1 Then CFG.BufferCount = 3

    'backbuffer
    If Option2.Value Then
        CFG.BK_FMT = DEV.CurrBKFMT
        CFG.DP_FMT = DEV.CurrentWINDOWED.DP_FMT(Combo2.ListIndex).FMT(Combo1.ListIndex)
        CFG.IS_FullScreen = 0
    End If
    If Option1.Value Then
        CFG.BK_FMT = DEV.RESO(NemoCmbRes.ListIndex).BK_FMT(Combo2.ListIndex)
        CFG.DP_FMT = DEV.RESO(NemoCmbRes.ListIndex).DP_FMT(Combo2.ListIndex).FMT(Combo1.ListIndex)
        CFG.IS_FullScreen = True
    End If

    If Not CFG.IS_FullScreen Then CFG.BufferCount = 1

    CFG.ForceVerSINC = Check1.Value

    CFG.Bpp = GetBKBit(CFG.BK_FMT)

    If Check1.Value Then CFG.ForceVerSINC = 0 Else CFG.ForceVerSINC = 1

    'acceleration
    If NemoCmbDevice.ListIndex = 0 Then CFG.DeviceTyp = D3DDEVTYPE_REF

    If NemoCmbDevice.ListIndex = 1 Then CFG.DeviceTyp = D3DDEVTYPE_HAL

    If NemoCmbDevice.ListIndex = 2 Then
        CFG.USE_TnL = True
        CFG.DeviceTyp = D3DDEVTYPE_HAL
    End If

    CFG.GamaLevel = GAMA_VAL

    'Call InitializeNemo(lpNEMO)
    'lpNEMO.Initialize CFG.appHandle, CFG.DeviceTyp, Not (CFG.IS_FullScreen), CFG.width, CFG.height, CFG.BPP, CFG.USE_TnL, Not CFG.ForceVerSINC, CFG.GamaLevel, CFG.BufferCount

    Unload Me

End Sub

Private Sub DoButton()

  Dim DC As Long

    DC = GetDC(Command1.hwnd)
    Ellipse GetDC(Command1.hwnd), 10, 10, 80, 31

End Sub

Private Sub EnumerateAdapters()

  Dim I As Integer, sTemp As String, J As Integer

  ''//This'll either be 1 or 2

    nAdapters = TempD3D8.GetAdapterCount

    For I = 0 To nAdapters - 1
        'Get the relevent Details
        TempD3D8.GetAdapterIdentifier I, 0, AdapterInfo

        'Get the name of the CurrentWINDOWED adapter - it's stored as a long
        'list of character codes that we need to parse into a string
        ' - Dont ask me why they did it like this; seems silly really :)
        sTemp = "" 'Reset the string ready for our use
        For J = 0 To 511
            sTemp = sTemp & Chr$(AdapterInfo.Description(J))
        Next J
        sTemp = Replace(sTemp, Chr$(0), " ")
        NemoCmbAdapters.AddItem sTemp
    Next I

End Sub

Private Sub EnumerateDevices()

    On Local Error Resume Next ''//We want to handle the errors...
      Dim CAPS As D3DCAPS8

        TempD3D8.GetDeviceCaps NemoCmbAdapters.ListIndex, D3DDEVTYPE_HAL, CAPS
        If Err.Number = D3DERR_NOTAVAILABLE Then
            'There is no hardware acceleration
            NemoCmbDevice.AddItem "Reference Rasterizer (REF)" 'Reference device will always be available
          Else
            NemoCmbDevice.AddItem "Hardware Acceleration (HAL)"
            NemoCmbDevice.AddItem "Reference Rasterizer (REF)" 'Reference device will always be available
        End If

End Sub

Sub EnumerateDispModes2()

  Dim I, K
  Dim DDM As D3DDISPLAYMODE
  Dim dd As D3DDISPLAYMODE

    K = TempD3D8.GetAdapterModeCount(0)
    TempD3D8.GetAdapterDisplayMode 0, dd

    For I = 0 To K - 1
        Call TempD3D8.EnumAdapterModes(0, I, DDM)
        FillFMT DDM
    Next I

    For I = 0 To DEV.NumRES - 1
        NemoCmbRes.AddItem " " + Str(DEV.RESO(I).Wi) + " x " + Str(DEV.RESO(I).HI) '+ STR(DEV.RESO(I).NumBK)
        If dd.Width = DEV.RESO(I).Wi And dd.Height = DEV.RESO(I).HI Then
            NemoCmbRes.ListIndex = I
            CurrINDEX = I

        End If

        'FillBackBuff DEV.RESO(I), DEV.CurrBKFMT, 0
        FillDepthBuff DEV.RESO(I), 0
    Next I

End Sub

Private Sub EnumerateHardware(Renderer As Long)

  'Renderer = Renderer + 1 ''//We need it on a base 1 scale (not base 0)
  'List1.Clear ''//Clear our list

  Dim CAPS As D3DCAPS8 ''//Holds all our information...

  '

    TempD3D8.GetDeviceCaps NemoCmbAdapters.ListIndex, D3DDEVTYPE_HAL, CAPS

    NemoCmbDevice.AddItem "HEL device (REF)"

    NemoCmbDevice.AddItem "Hardware Abstaction Layer (HAL)"

    If CAPS.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT Then

        If NemoCmbDevice.ListCount < 3 Then _
           NemoCmbDevice.AddItem "Hardware Transform and lighting (TnL)"

    End If

    '
    If CAPS.Caps2 And D3DCAPS2_FULLSCREENGAMMA Then
        'Frame1.Enabled = True
    End If

    NemoCmbDevice.ListIndex = 2

End Sub

Sub FillWindowedList()

  Dim I, K

    Combo1.Clear
    Combo2.Clear
    If DEV.CurrBKFMT = 0 Then Exit Sub

    For I = 0 To DEV.CurrentWINDOWED.NumBK - 1

        Combo2.AddItem FMTsTR(DEV.CurrentWINDOWED.BK_FMT(I)) + " (" + FMTsTR2(DEV.CurrentWINDOWED.BK_FMT(I)) + ")"
        If DEV.CurrentWINDOWED.BK_FMT(I) = DEV.CurrBKFMT Then K = I
    Next I
    'If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
    Combo2.ListIndex = K
    Combo1.Clear
    For I = 0 To DEV.CurrentWINDOWED.DP_FMT(K).NumD - 1
        Combo1.AddItem FMTsTR(DEV.CurrentWINDOWED.DP_FMT(K).FMT(I))

    Next I

    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0

End Sub

Private Sub Form_Load()

'    If IS_GLOB Then Exit Sub
'  Dim N As New NemoX

    GAMA_VAL = 1

   
    Init
    'Frame1.Enabled = 0
    ''//1. Create any relevent objects
    Set TempDX8 = New DirectX8
    Set TempD3D8 = TempDX8.Direct3DCreate

    EnumerateAdapters
    NemoCmbAdapters.ListIndex = NemoCmbAdapters.ListCount - 1

    CFG.IS_FullScreen = False
    Option2.Value = True
    Option1.Value = 0

    FillCurrentWINDOWED
    FillBackBuff DEV.CurrentWINDOWED, DEV.CurrBKFMT, True
    FillDepthBuff DEV.CurrentWINDOWED, 0
    FillWindowedList

    EnumerateHardware 0
    EnumerateDispModes2

    Command2.Refresh

    keybd_event VK_RIGHT, 0, 0, 0   ' press H
    keybd_event VK_RIGHT, 0, KEYEVENTF_KEYUP, 0   ' release H

    SetFocu Command1.hwnd
    DoButton

End Sub

Public Sub Free()

    Hide

    Unload Me

End Sub

Private Sub GAMdown_Click(Index As Integer)

    GAMA_VAL = GAMA_VAL - 0.05
    Text1.Text = Round(GAMA_VAL, 2)

End Sub

Private Sub GAMup_Click(Index As Integer)

    GAMA_VAL = GAMA_VAL + 0.05
    Text1.Text = Round(GAMA_VAL, 2)

End Sub

Private Sub Init()

    With CFG
        CFG.Bpp = 16
        .DeviceTyp = D3DDEVTYPE_HAL
        .ForceVerSINC = True
        .GamaLevel = 1
        .BufferCount = 1

    End With

    Text1.Text = Str(CFG.GamaLevel)
    Option1.Value = 1
    NemoCmbRes.Enabled = 1
    Check1.Value = 1
    Option3.Value = 1

End Sub

Private Sub NemoCmbDevice_Click()

    If NemoCmbDevice.ListIndex = 0 Then CFG.DeviceTyp = D3DDEVTYPE_REF

    If NemoCmbDevice.ListIndex = 1 Then CFG.DeviceTyp = D3DDEVTYPE_HAL

    If NemoCmbDevice.ListIndex = 2 Then CFG.USE_TnL = True

End Sub

Private Sub NemoCmbRes_Click()

  Dim V1, V2, SS As String

    SS = NemoCmbRes.List(NemoCmbRes.ListIndex)

    V1 = InStr(SS, "x")
    CFG.Width = Val(Left$(SS, V1 - 1))

    V2 = InStr(V1, SS, "x")
    CFG.Height = Val(Right$(SS, Len(SS) - V2))

End Sub

'Private Sub HScroll1_Change()
'
'    CFG.GamaLevel = (HScroll1 / 100) * 2
'    Text1.Text = STR(CFG.GamaLevel)
'
'End Sub

Private Sub Option1_Click()

    Option1.Value = True
    Option2.Value = 0

    Set_Fmode
    Combo2.Enabled = 1

End Sub

Private Sub Option2_Click()

    Option1.Value = 0
    Option2.Value = 1

    Set_Fmode
    Combo2.Enabled = False
    FillWindowedList
    'Combo2_Click

End Sub

Private Sub Option3_Click()

    Option3.Value = True
    Option4.Value = 0

    If Option3.Value = 1 Then CFG.BufferCount = 2

End Sub

Private Sub Option4_Click()

    Option3.Value = 0
    Option4.Value = 1

    If Option4.Value = 1 Then CFG.BufferCount = 3

End Sub

Sub Set_Fmode()

    CFG.IS_FullScreen = Option1.Value
    'Check1.Enabled = Int(Option1.value)
    'Frame1.Enabled = Option1.value
    'HScroll1.Enabled = Option1.value
    Text1.Enabled = Option1.Value
    GAMup(0).Enabled = Option1.Value
    GAMdown(0).Enabled = Option1.Value

    NemoCmbRes.Enabled = Option1.Value

End Sub

Sub SHOW_DIALOG(LpHandle As Long)

    Load Me
    CFG.appHandle = LpHandle
    
    Show vbModal

End Sub








Private Sub FillBackBuff(CUR As tFMT, DispFMT As Long, Optional ByVal winDowed As Boolean = False)

  Dim BB(5) As CONST_D3DFORMAT
  Dim I, J

    BB(0) = D3DFMT_R5G6B5
    'BB(5) = D3DFMT_A2R10G10B10
    BB(4) = D3DFMT_A8R8G8B8
    BB(3) = D3DFMT_X8R8G8B8
    BB(2) = D3DFMT_A1R5G5B5
    BB(1) = D3DFMT_X1R5G5B5

    For I = 0 To 4

        If TempD3D8.CheckDeviceType(0, D3DDEVTYPE_HAL, BB(I), BB(I), winDowed) = D3D_OK Then

            CUR.NumBK = CUR.NumBK + 1
            ReDim Preserve CUR.BK_FMT(CUR.NumBK - 1)
            CUR.BK_FMT(CUR.NumBK - 1) = BB(I)

        End If

    Next I

End Sub

'====================ENUMARATION========

Private Sub FillCurrentWINDOWED()

  Dim DM As D3DDISPLAYMODE

    TempD3D8.GetAdapterDisplayMode 0, DM

    DEV.CurrentWINDOWED.HI = DM.Height
    DEV.CurrentWINDOWED.Wi = DM.Width
    DEV.CurrBKFMT = DM.format

End Sub

Private Sub FillDepthBuff(CUR As tFMT, DispFMT As Long)

  Dim BB(5) As CONST_D3DFORMAT
  Dim I, J, K

    BB(0) = D3DFMT_D16
    BB(1) = D3DFMT_D15S1
    BB(2) = D3DFMT_D24X8
    BB(3) = D3DFMT_D24S8
    BB(4) = D3DFMT_D24X4S4
    BB(5) = D3DFMT_D32

    ReDim CUR.DP_FMT(CUR.NumBK - 1)
    For J = 0 To CUR.NumBK - 1

        For I = 0 To 5
            If TempD3D8.CheckDeviceFormat(0, D3DDEVTYPE_HAL, CUR.BK_FMT(J), D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, BB(I)) = D3D_OK Then

                If TempD3D8.CheckDepthStencilMatch(0, D3DDEVTYPE_HAL, CUR.BK_FMT(J), CUR.BK_FMT(J), BB(I)) = D3D_OK Then

                    CUR.DP_FMT(J).NumD = CUR.DP_FMT(J).NumD + 1
                    ReDim Preserve CUR.DP_FMT(J).FMT(CUR.DP_FMT(J).NumD - 1)

                    CUR.DP_FMT(J).FMT(CUR.DP_FMT(J).NumD - 1) = BB(I)
                End If

            End If

        Next I

    Next J

End Sub

Private Sub FillFMT(DM As D3DDISPLAYMODE)

  Dim I, J, K

    If DEV.NumRES = 0 Then
        DEV.NumRES = 1
        ReDim DEV.RESO(DEV.NumRES - 1)
        DEV.RESO(DEV.NumRES - 1).Wi = DM.Width
        DEV.RESO(DEV.NumRES - 1).HI = DM.Height
        DEV.RESO(DEV.NumRES - 1).NumBK = DEV.RESO(DEV.NumRES - 1).NumBK + 1
        ReDim Preserve DEV.RESO(DEV.NumRES - 1).BK_FMT(DEV.RESO(DEV.NumRES - 1).NumBK - 1)
        DEV.RESO(DEV.NumRES - 1).BK_FMT(DEV.RESO(DEV.NumRES - 1).NumBK - 1) = DM.format

        Exit Sub
    End If

    For I = 0 To DEV.NumRES - 1

        If DEV.RESO(I).HI = DM.Height And DEV.RESO(I).Wi = DM.Width And DEV.RESO(I).NumBK > 0 Then

            K = DEV.RESO(I).NumBK
            If DEV.RESO(I).BK_FMT(K - 1) <> DM.format Then
                DEV.RESO(I).NumBK = DEV.RESO(I).NumBK + 1
                ReDim Preserve DEV.RESO(I).BK_FMT(DEV.RESO(I).NumBK - 1)
                DEV.RESO(I).BK_FMT(DEV.RESO(I).NumBK - 1) = DM.format
                Exit Sub
            End If

            Exit Sub
        End If

    Next I

    DEV.NumRES = DEV.NumRES + 1
    ReDim Preserve DEV.RESO(DEV.NumRES - 1)
    DEV.RESO(DEV.NumRES - 1).Wi = DM.Width
    DEV.RESO(DEV.NumRES - 1).HI = DM.Height
    DEV.RESO(DEV.NumRES - 1).NumBK = DEV.RESO(DEV.NumRES - 1).NumBK + 1
    ReDim Preserve DEV.RESO(DEV.NumRES - 1).BK_FMT(DEV.RESO(DEV.NumRES - 1).NumBK - 1)
    DEV.RESO(DEV.NumRES - 1).BK_FMT(DEV.RESO(DEV.NumRES - 1).NumBK - 1) = DM.format

End Sub




Private Function FMTsTR(FM As Long) As String

    Select Case FM
      Case D3DFMT_UNKNOWN
        FMTsTR = "Surface format is unknown."
      Case D3DFMT_R8G8B8
        FMTsTR = "24-bit RGB pixel format."
      Case D3DFMT_A8R8G8B8
        FMTsTR = "32-bit ARGB pixel format with alpha."
      Case D3DFMT_X8R8G8B8
        FMTsTR = "32-bit" ' RGB"
      Case D3DFMT_R5G6B5
        FMTsTR = "16-bit" ' RGB"
      Case D3DFMT_X1R5G5B5
        FMTsTR = "16-bit pixel format where 5 bits are reserved for each color."
      Case D3DFMT_A1R5G5B5
        FMTsTR = "16-bit pixel format where 5 bits are reserved for each color and 1 bit is reserved for alpha (transparent texel)."
      Case D3DFMT_A4R4G4B4
        FMTsTR = "16-bit ARGB pixel format."
      Case D3DFMT_R3G3B2
        FMTsTR = "8-bit RGB texture format."
      Case D3DFMT_A8
        FMTsTR = "8-bit alpha only."
      Case D3DFMT_A8R3G3B2
        FMTsTR = "16-bit ARGB texture format."
      Case D3DFMT_X4R4G4B4
        FMTsTR = "16-bit RGB pixel format where 4 bits are reserved for each color."
      Case D3DFMT_A8P8
        FMTsTR = "8-bit color indexed with 8 bits of alpha."
      Case D3DFMT_P8
        FMTsTR = "8-bit color indexed."
      Case D3DFMT_L8
        FMTsTR = "8-bit luminance only."
      Case D3DFMT_A8L8
        FMTsTR = "16-bit alpha luminance."
      Case D3DFMT_A4L4
        FMTsTR = "8-bit alpha luminance."
      Case D3DFMT_V8U8
        FMTsTR = "16-bit bump-map format."
      Case D3DFMT_L6V5U5
        FMTsTR = "16-bit bump-map format with luminance."
      Case D3DFMT_X8L8V8U8
        FMTsTR = "32-bit bump-map format with luminance where 8 bits are reserved for each element."
      Case D3DFMT_Q8W8V8U8
        FMTsTR = "32-bit bump-map format."
      Case D3DFMT_V16U16
        FMTsTR = "32-bit bump-map format."
      Case D3DFMT_W11V11U10
        FMTsTR = "32-bit bump-map format."

      Case D3DFMT_D16_LOCKABLE
        FMTsTR = "16-bit z-buffer bit depth. This is an application-lockable surface format."
      Case D3DFMT_D32
        FMTsTR = "32-bit z-buffer bit depth."
      Case D3DFMT_D15S1
        FMTsTR = "16-bit z-buffer/1 bit stencil channel."
      Case D3DFMT_D24S8
        FMTsTR = "32-bit z-buffer 24/8 stencil channel."
      Case D3DFMT_D16
        FMTsTR = "16-bit z-buffer bit depth."
      Case D3DFMT_D24X8
        FMTsTR = "32-bit z-buffer bit depth."
      Case D3DFMT_D24X4S4
        FMTsTR = "32-bit z-buffer 24/4 bits stencil channel."
      Case D3DFMT_VERTEXDATA

        'Describes a vertex buffer surface.
      Case D3DFMT_INDEX16
        FMTsTR = "16-bit index buffer bit depth."
      Case D3DFMT_INDEX32
        FMTsTR = "32-bit index buffer bit depth."

    End Select

End Function

Private Function FMTsTR2(FMT As Long) As String

    Select Case FMT
      Case D3DFMT_UNKNOWN
        FMTsTR2 = "UNKNOWN"

      Case D3DFMT_R8G8B8
        FMTsTR2 = "R8G8B8"
      Case D3DFMT_A8R8G8B8
        FMTsTR2 = "A8R8G8B8"
      Case D3DFMT_X8R8G8B8
        FMTsTR2 = "X8R8G8B8"
      Case D3DFMT_R5G6B5
        FMTsTR2 = "R5G6B5"
      Case D3DFMT_X1R5G5B5
        FMTsTR2 = "X1R5G5B5"
      Case D3DFMT_A1R5G5B5
        FMTsTR2 = "A1R5G5B5"
      Case D3DFMT_A4R4G4B4
        FMTsTR2 = "A4R4G4B4"
      Case D3DFMT_R3G3B2
        FMTsTR2 = "R3G3B2"
      Case D3DFMT_A8
        FMTsTR2 = "A8"
      Case D3DFMT_A8R3G3B2
        FMTsTR2 = "A8R3G3B2"
      Case D3DFMT_X4R4G4B4
        FMTsTR2 = "X4R4G4B4"

      Case D3DFMT_A8P8
        FMTsTR2 = "A8P8"
      Case D3DFMT_P8
        FMTsTR2 = "P8"

      Case D3DFMT_L8
        FMTsTR2 = "L8"
      Case D3DFMT_A8L8
        FMTsTR2 = "A8L8"
      Case D3DFMT_A4L4
        FMTsTR2 = "A4L4"

      Case D3DFMT_V8U8
        FMTsTR2 = "V8U8"
      Case D3DFMT_L6V5U5
        FMTsTR2 = "L6V5U5"
      Case D3DFMT_X8L8V8U8
        FMTsTR2 = "X8L8V8U8"
      Case D3DFMT_Q8W8V8U8
        FMTsTR2 = "Q8W8V8U8"
      Case D3DFMT_V16U16
        FMTsTR2 = "V16U16"

      Case D3DFMT_UYVY
        FMTsTR2 = "UYVY"
      Case D3DFMT_YUY2
        FMTsTR2 = "YUY2"
      Case D3DFMT_DXT1
        FMTsTR2 = "DXT1"
      Case D3DFMT_DXT2
        FMTsTR2 = "DXT2"
      Case D3DFMT_DXT3
        FMTsTR2 = "DXT3"
      Case D3DFMT_DXT4
        FMTsTR2 = "DXT4"
      Case D3DFMT_DXT5
        FMTsTR2 = "DXT5"

      Case D3DFMT_D16_LOCKABLE
        FMTsTR2 = "D16_LOCKABLE"
      Case D3DFMT_D32
        FMTsTR2 = "D32"
      Case D3DFMT_D15S1
        FMTsTR2 = "D15S1"
      Case D3DFMT_D24S8
        FMTsTR2 = "D24S8"
      Case D3DFMT_D16
        FMTsTR2 = "D16"
      Case D3DFMT_D24X8
        FMTsTR2 = "D24X8"
      Case D3DFMT_D24X4S4
        FMTsTR2 = "D24X4S4"

      Case D3DFMT_VERTEXDATA
        FMTsTR2 = "VERTEXDATA"
      Case D3DFMT_INDEX16
        FMTsTR2 = "INDEX16"
      Case D3DFMT_INDEX32
        FMTsTR2 = "INDEX32"

    End Select

End Function



Function GetBKBit(FMT As Long) As Long

    Select Case FMT

        '// 32 bit modes
      Case D3DFMT_A8R8G8B8
        GetBKBit = 32

      Case D3DFMT_X8R8G8B8
        GetBKBit = 32

        '// 24 bit modes
      Case D3DFMT_R8G8B8
        GetBKBit = 24

        '// 16 bit modes
      Case D3DFMT_R5G6B5
        GetBKBit = 16
      Case D3DFMT_X1R5G5B5
        GetBKBit = 16
      Case D3DFMT_A1R5G5B5
        GetBKBit = 16
      Case D3DFMT_A4R4G4B4
        GetBKBit = 16

    End Select

End Function

Function GetDepthBits(FMT As Long) As Long

    Select Case FMT

      Case D3DFMT_D16
        GetDepthBits = 16
      Case D3DFMT_D16_LOCKABLE
        GetDepthBits = 16

      Case D3DFMT_D15S1
        GetDepthBits = 15

      Case D3DFMT_D24X8
        GetDepthBits = 24
      Case D3DFMT_D24S8
        GetDepthBits = 24
      Case D3DFMT_D24X4S4
        GetDepthBits = 24

      Case D3DFMT_D32
        GetDepthBits = 32

    End Select

End Function





