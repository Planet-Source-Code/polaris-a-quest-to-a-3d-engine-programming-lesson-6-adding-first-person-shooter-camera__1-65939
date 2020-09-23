VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuest3D_Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================================================
'
'       THIS IS THE CORE CLASS OF THE ENGINE
'
'  - we start 3D and close 3D here
'  - We acces the GFX here
'  - we change all GFX state here
'================================================================================================

'Lesson 1
'
'We write
'  - an initialization method
'  - Begin3D to start 3d rendering
'  - End3D to close 3D drawing and flip all draw stuff to the screen
'  - Free to close engine
'  - GetKeyPressed() to get the current pressed key

'=====================================================
'Method:Init
' Create 3D device and all objects needed
' - Windowed for create a windowed device
' - FullScreenWidth and FullScreenheight screen size
' - DisableVerticalSync to disable monitor vertical refresh
'    if DisableVerticalSync=true the frames per second increase hugely

'Return: Boolean
' return true if the engine creation suceed
' return false if there was an error
'=====================================================

Enum QUEST3D_FILTERING_MODE
    QUEST3D_FILTER_POINT = 0
    QUEST3D_FILTER_BILINEAR = 1
    QUEST3D_FILTER_TRILINEAR = 2
    QUEST3D_FILTER_ANISOTROPIC = 3
    QUEST3D_FILTER_FLATCUBIC = 4
    QUEST3D_FILTER_GAUSSIANCUBIC = 5

End Enum

Enum QUEST3D_SHADING_MODE
    QUEST3D_SHADE_FLAT = D3DSHADE_FLAT
    QUEST3D_SHADE_GOURAUD = D3DSHADE_GOURAUD
    QUEST3D_SHADE_PHONG = D3DSHADE_PHONG

End Enum

Enum QUEST3D_FILL_MODE
    QUEST3D_FILL_WIREFRAME = 0
    QUEST3D_FILL_SOLID = 1

End Enum

Function Init(ByVal FormHandle As Long, Optional ByVal winDowed As Boolean = 1, Optional ByVal FullScreenWidth As Integer = 640, Optional ByVal FullScreenHeight As Integer = 480, Optional ByVal DisableVerticalSync As Boolean = True) As Boolean

  Dim RectClient As RECT

    Set obj_D3D = obj_DX.Direct3DCreate() 'we create the D3D object

    Set obj_D3DX = New D3DX8

    ' Get The current Display Mode format
  Dim mode As D3DDISPLAYMODE
    obj_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode

    'we set the parameters for the creation of the Device

    If winDowed Then
        'for Windowed mode
        Data.Hwindow = FormHandle
        Data.WinParam.winDowed = 1
        Data.WinParam.BackBufferCount = 1 '1 backbuffer
        Data.WinParam.BackBufferFormat = mode.format 'color format
        Data.WinParam.SwapEffect = D3DSWAPEFFECT_FLIP

        If DisableVerticalSync = False Then Data.WinParam.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC

      Else
        'for fullscreen

        Data.Buffer_Rect.Left = 0
        Data.Buffer_Rect.Top = 0
        Data.Buffer_Rect.Right = FullScreenWidth
        Data.Buffer_Rect.bottom = FullScreenHeight

        Data.WinParam.BackBufferCount = 1
        Data.WinParam.BackBufferFormat = mode.format
        Data.WinParam.BackBufferWidth = FullScreenWidth
        Data.WinParam.BackBufferHeight = FullScreenHeight
        Data.WinParam.SwapEffect = D3DSWAPEFFECT_COPY

        If Not DisableVerticalSync = False Then
            Data.WinParam.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
          Else 'VERTSINCOFF = FALSE
            Data.WinParam.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
        End If
    End If

    Data.WinParam.EnableAutoDepthStencil = 1
    Data.WinParam.AutoDepthStencilFormat = D3DFMT_D16 '16 bit Z-Buffer
    Data.WinParam.hDeviceWindow = FormHandle  'target

    GetWindowRect Data.Hwindow, RectClient
    CFG.Width = RectClient.Right - RectClient.Left
    CFG.Height = RectClient.bottom - RectClient.Top
    Data.Buffer_Rect.Right = CFG.Width
    Data.Buffer_Rect.bottom = CFG.Height

    Data.Buffer_Width = CFG.Width
    Data.Buffer_Height = CFG.Height

    'create the device
    'D3DADAPTER_DEFAULT means we use default Graphic card VGA is the first
    'D3DCREATE_SOFTWARE_VERTEXPROCESSING we use directx driver software vertex processing

    Set obj_Device = obj_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, FormHandle, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Data.WinParam)

    'if the device is nothing we fail else Great WE DID IT nice!!
    Data.Is_engineActive = Not (obj_Device Is Nothing)

    'we prepare informations
    If Data.Is_engineActive Then
        DoDefault
    End If

    Init = Data.Is_engineActive

End Function

Function Init_Dialogue(ByVal appHandle As Long) As Boolean

    frmEnum.SHOW_DIALOG appHandle

    If IS_ERROR Then
        Init_Dialogue = False
        Exit Function

    End If

    Set obj_DX = New DirectX8
    Set obj_D3DX = New D3DX8
    Set obj_D3D = obj_DX.Direct3DCreate() ' Create D3D

  Dim DispMode As D3DDISPLAYMODE 'for current displaymode
  Dim D3DWindow As D3DPRESENT_PARAMETERS 'for present parameters
  Dim RectClient As RECT

    GetWindowRect CFG.appHandle, RectClient
    If Not CFG.IS_FullScreen Then
        'inizializza finestra
        D3DWindow.winDowed = 1
        D3DWindow.BackBufferCount = 1 '1 backbuffer
        D3DWindow.BackBufferFormat = CFG.BK_FMT 'color format from dialogue
        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP

        If CFG.ForceVerSINC Then D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC

        GetWindowRect CFG.appHandle, RectClient
        CFG.Width = RectClient.Right - RectClient.Left
        CFG.Height = RectClient.bottom - RectClient.Top
        Data.Buffer_Rect.Right = CFG.Width
        Data.Buffer_Rect.bottom = CFG.Height
      Else
        'fullscreen
        'GetWindowRect CFG.appHandle, Data.Buffer_Rect

        Data.Buffer_Rect.Left = 0
        Data.Buffer_Rect.Top = 0
        Data.Buffer_Rect.Right = CFG.Width
        Data.Buffer_Rect.bottom = CFG.Height

        D3DWindow.BackBufferCount = CFG.BufferCount  ' backbuffer
        D3DWindow.BackBufferFormat = CFG.BK_FMT 'colore
        D3DWindow.BackBufferWidth = CFG.Width
        D3DWindow.BackBufferHeight = CFG.Height
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY

        If Not CFG.ForceVerSINC Then
            D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
          Else 'VERTSINCOFF = FALSE
            D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
        End If
    End If
    'comuni
    D3DWindow.EnableAutoDepthStencil = 1
    D3DWindow.AutoDepthStencilFormat = CFG.DP_FMT 'CFG.BK_FMT '16 bit Z-Buffer
    D3DWindow.hDeviceWindow = CFG.appHandle  'target
    'If debugMode Then D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    'If (GetStencilBits(CFG.DP_FMT) > 0) Then _
        'D3DWindow.Flags = D3DPRESENTFLAG_DISCARD_DEPTHSTENCIL

    'crea device

    If CFG.DeviceTyp = D3DDEVTYPE_HAL Then
        Set obj_Device = obj_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, CFG.appHandle, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
      ElseIf CFG.USE_TnL Then
        Set obj_Device = obj_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, CFG.appHandle, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
      Else
        Set obj_Device = obj_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, CFG.appHandle, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)

    End If

    Data.Buffer_Width = CFG.Width
    Data.Buffer_Height = CFG.Height
    Data.Bpp = CFG.Bpp
    Data.Is_Windowed = Not CFG.IS_FullScreen
    Data.GamaLevel = CFG.GamaLevel

    Data.WinParam = D3DWindow

    Init_Dialogue = Not (obj_Device Is Nothing)

    If Init_Dialogue Then
        Data.Is_engineActive = True
        DoDefault
    End If

End Function

Private Sub DoDefault()

  'here we set default parameters

  'we set the color to blue

    Data.BackBuff_ClearColor = D3DColorXRGB(0, 0, 255)

    Me.Set_Font "arial", 12

    Me.Set_ViewFrustum 5, 2500, QUEST3D_PI / 4, 1

    'this is for texture filtering
    Data.TEXTURE_FILTER = D3DX_FILTER_LINEAR
    Data.TEXTURE_MIPMAPFILTER = QUEST3D_TEXTURE_DEFAULT

    'we set default material color
  Dim COL As D3DCOLORVALUE
  Dim mtrl As D3DMATERIAL8
    With COL

        .R = 1
        .g = 1
        .b = 1
        .A = 1

    End With
    mtrl.diffuse = COL
    mtrl.Ambient = COL
    With COL

        .R = 0.09
        .g = 0.09
        .b = 0.09
        .A = 0.01

    End With

    mtrl.specular = COL
    Data.Init_Material = mtrl

    obj_Device.SetMaterial mtrl
    Set_EngineAmbientColor &HFFFFFFFF

    Init_State

    'we enable zbuffer
    obj_Device.SetRenderState D3DRS_ZENABLE, 1
    'we activate zbuffer and back_buffer clearing
    Data.IS_ClearRenderTarget = True

    Get_Capabilities

    'aasign a global pointer to that interface
    'so we will able to use anywhere the methods inside
    'that class
    
    Dim ParentHandle As Long
    
    ParentHandle = GetParent(Data.Hwindow)
    
   

    If ParentHandle <> 0 Then
        Call Init_Input(ParentHandle)
      Else
        Call Init_Input(Data.Hwindow)
    End If
    
     If ParentHandle = 0 Then ParentHandle = Data.Hwindow

    Data.HwindowParent = ParentHandle

    
    Set LpGLOBAL_QUEST3D = Me

End Sub

Private Sub Init_State()

  Dim I As Long
  Dim J As Integer
  Dim NumOfStage As Integer

    NumOfStage = 2

    For I = 0 To NumOfStage - 1
        For J = 1 To 28
            Data.lpState.Init_TexSate(I, J) = obj_Device.GetTextureStageState(I, J)
        Next J
    Next I

    I = 7
    While I <= 171

        If I <> 10 And I <> 11 And I <> 12 And I <> 13 And I <> 17 And I <> 18 And _
           I <> 21 And ((I < 31) Or (I > 39)) And ((I < 41) Or (I > 46)) And _
           ((I < 49) Or (I > 51)) And ((I < 61) Or (I > 127)) And I <> 144 And I <> 149 And I <> 150 And I <> 169 Then _
           Data.lpState.Init_Renderstate(I) = obj_Device.GetRenderState(I)
        I = I + 1
    Wend

    Me.Set_EngineShadeMode QUEST3D_SHADE_GOURAUD

End Sub

Sub FreeEngine()

  'free all textures

  Dim I As Long

    For I = 0 To myTEXPOOL.NumTextureInpool - 1
        Set myTEXPOOL.POOL_texture(I) = Nothing

    Next I

    'free inputs

    If Data.IS_DinputOK Then
        'free input device
        Set obj_Dinput = Nothing

        'On Error Resume Next
        DIKeyBoardDevice.Unacquire

        DIMouseDevice.Unacquire

        If Data.IS_Joystick Then

            For I = 0 To Data.JoyNumDevice - 1
                DIjoyDevice(I).Unacquire
                Set DIjoyDevice(I) = Nothing

            Next I
            Erase DIjoyDevice
            Erase AxisPresent

        End If

        Set DIKeyBoardDevice = Nothing
        Set DIMouseDevice = Nothing
    End If

    'we deallocate memory and free objects
    Set obj_DX = Nothing
    Set obj_D3D = Nothing
    Set obj_Device = Nothing

    Data.Is_engineActive = False

End Sub

'=======================================
'We tell D3D that we Begin 3D rendering
'
'======================================
Sub Begin3D()

    Data.Total_TriangleRENDERED = 0
    Data.Total_VerticeRENDERED = 0
    ComputeFrames

    If Data.Is_engineActive = False Then Exit Sub

    'we clear previous drawn
    If Data.IS_ClearRenderTarget Then
        obj_Device.Clear ByVal 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Data.BackBuff_ClearColor, 1, 0
      Else
        obj_Device.Clear ByVal 0, ByVal 0, D3DCLEAR_ZBUFFER, Data.BackBuff_ClearColor, 1, 0

    End If

    'We tell the device we are ready to render triangles
    obj_Device.BeginScene

End Sub

'==========================================================
'Here we compute frame per second
'
'==========================================================
Private Sub ComputeFrames()

  Dim Elapsed As Single

    'we get the current time in second
    Data.Fps_TimePassed = timeGetTime * 0.001 - Data.Fps_CurrentTime
    Data.Fps_CurrentTime = timeGetTime * 0.001
    'we count Frames
    Data.Fps_FrameCounter = Data.Fps_FrameCounter + 1

    'we compute time passed during two rendering scene
    Elapsed = Data.Fps_CurrentTime - Data.Fps_LastTime

    'if we are over 1 sec
    If Elapsed > 1 Then
        'we reset the last time to Now
        Data.Fps_LastTime = Data.Fps_CurrentTime

        'we check that we have 1 second,1.004=85 is not true so we do (85 x 1)/1.004
        If Elapsed <> 0 Then _
           Data.Fps_FramePerSecond = Data.Fps_FrameCounter / (Elapsed)
        'we reset the counter
        Data.Fps_FrameCounter = 0

    End If

End Sub

Function Get_NumberOfPolygonDrawn() As Long

    Get_NumberOfPolygonDrawn = Data.Total_TriangleRENDERED

End Function

Function Get_NumberOfVerticeDrawn() As Long

    Get_NumberOfVerticeDrawn = Data.Total_VerticeRENDERED

End Function

Function Get_D3DDevice() As Direct3DDevice8

    Set Get_D3DDevice = obj_Device

End Function

Function Get_FramesPerSeconde() As Single

    Get_FramesPerSeconde = Data.Fps_FramePerSecond

End Function

Function Get_TimePassed() As Single

    Get_TimePassed = Data.Fps_TimePassed

End Function

'=========================================
'We tell D3D that we Stop 3D rendering
'=========================================
Sub End3D()

    If Data.Is_engineActive = False Then Exit Sub

    obj_Device.EndScene 'We tell the device we'are now finished
    obj_Device.Present ByVal 0, ByVal 0, 0, ByVal 0 'we flip all drawn stuffs over the monitor' screen

End Sub

Function Get_KeyPressedApi(ByVal KEY As KeyCodeConstants) As Boolean

    Get_KeyPressedApi = (GetAsyncKeyState(KEY) <> 0)

End Function

Function Get_KeyPress(ByVal KEY As QUEST3D_KEY_CONST) As Boolean

    If Data.IS_DinputOK Then

        DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE

        ' If key is pressed then
        If DIKEYBOARDSTATE.KEY(KEY) <> 0 Then Get_KeyPress = True

    End If

End Function

Private Sub Class_Initialize()

    Randomize Timer

End Sub

Sub Set_Font(ByVal FontName As String, ByVal PoliceSize As Integer)

    Set Data.StFont = New StdFont

    Data.StFont.Name = FontName
    Data.StFont.Size = PoliceSize
    Set Data.FontDesc = Data.StFont

    Set Data.MainFont = obj_D3DX.CreateFont(obj_Device, Data.FontDesc.hFont)

End Sub

Sub Draw_Text(ByVal St As String, ByVal x As Long, ByVal y As Long, Optional ByVal color As Long = &HFFFFCC00)

  Dim TextRect As RECT

    TextRect.Top = y

    TextRect.Left = x

    obj_D3DX.DrawText Data.MainFont, color, St, TextRect, DT_TOP Or DT_LEFT

End Sub

'============================================================
'Engine State
'
'============================================================
Sub Set_BackbufferClearColor(ByVal ClearColor As Long)

    Data.BackBuff_ClearColor = ClearColor

End Sub

Sub Set_EngineVertexShader(ByVal FVF_format As Long)

    Call obj_Device.SetVertexShader(FVF_format)

End Sub

Sub Set_EngineLight(ByVal Enable As Boolean)

    obj_Device.SetRenderState D3DRS_LIGHTING, Enable

End Sub

Sub Set_EngineCullMode(ByVal CullMode As CONST_D3DCULL)

    obj_Device.SetRenderState D3DRS_CULLMODE, CullMode

End Sub

Sub Set_EngineClearRenderTarget(ByVal Enable As Boolean)

    Data.IS_ClearRenderTarget = Enable

End Sub

'----------------------------------------
'Name: Set_CullMode
'Object: Set
'Event: CullMode
'----------------------------------------
'----------------------------------------
'Name: Set_CullMode
'Object: Set
'Event: CullMode
'Description:
'----------------------------------------

Sub Set_EngineDitherEnable(ByVal Enable As Boolean)

    obj_Device.SetRenderState D3DRS_DITHERENABLE, Enable

End Sub

'----------------------------------------
'Name: Set_EngineFillMode
'Object: Set
'Event: EngineFillMode
'----------------------------------------
'----------------------------------------
'Name: Set_EngineFillMode
'Object: Set
'Event: EngineFillMode
'Description:
'----------------------------------------
Sub Set_EngineFillMode(ByVal FillMode As QUEST3D_FILL_MODE)

    If FillMode = QUEST3D_FILL_WIREFRAME Then
        Call obj_Device.SetRenderState(D3DRS_FILLMODE, D3DFILL_WIREFRAME)
        Call obj_Device.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        'Call obj_device.SetTexture(0, Nothing)

      Else
        Call obj_Device.SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call obj_Device.SetRenderState(D3DRS_CULLMODE, D3DCULL_CCW)
    End If

End Sub

'=======
' Engine Renderstate Parameters
'
'===

'Dim fBias As Single
'fBias = -0.5
'obj_device.SetTextureStageState 0, D3DTSS_MIPMAPLODBIAS, FloatToDWord(fBias)
'obj_device.SetTextureStageState 1, D3DTSS_MIPMAPLODBIAS, FloatToDWord(fBias)

Sub Set_EngineMipMapLodBias(ByVal BiasVal As Single)

  Dim Count As Long

    For Count = 0 To Data.Capa.Max_TextureStages - 1

        obj_Device.SetTextureStageState Count, D3DTSS_MIPMAPLODBIAS, FloatToDWord(BiasVal)
    Next Count

End Sub

Sub Set_EnginePixelShader(ByVal PixelShader As Long)

    Call obj_Device.SetPixelShader(PixelShader)

End Sub

Sub Set_EngineTextureFromPool(ByVal TextID As Integer, Optional ByVal TextureStage As Long = 0)

    If TextID < 0 Then Exit Sub
    obj_Device.SetTexture TextureStage, myTEXPOOL.POOL_texture(TextID)

End Sub

'----------------------------------------
'Name: Set_EngineRenderState
'Object: Set
'Event: EngineRenderState
'----------------------------------------
'----------------------------------------
'Name: Set_EngineRenderState
'Object: Set
'Event: EngineRenderState
'Description:
'----------------------------------------
Sub Set_EngineRenderState(ByVal RenderState As CONST_D3DRENDERSTATETYPE, Val As Long)

    obj_Device.SetRenderState RenderState, Val

End Sub

'----------------------------------------
'Name: Set_EngineShadeMode
'Object: Set
'Event: EngineShadeMode
'----------------------------------------
'----------------------------------------
'Name: Set_EngineShadeMode
'Object: Set
'Event: EngineShadeMode
'Description:
'----------------------------------------
Sub Set_EngineShadeMode(ByVal ShadingModel As QUEST3D_SHADING_MODE)

    Call obj_Device.SetRenderState(D3DRS_SHADEMODE, ShadingModel)

End Sub

Sub Set_EngineSpecularEnable(ByVal Enable As Boolean)

    obj_Device.SetRenderState D3DRS_SPECULARENABLE, Enable

End Sub

Sub Set_EngineTexture(BaseTexture As Direct3DBaseTexture8, Optional ByVal Stage As Long = 0)

    obj_Device.SetTexture Stage, BaseTexture

End Sub

Sub Set_EngineTextureCubic(Texture As Direct3DCubeTexture8, Optional ByVal Stage As Long = 0)

    obj_Device.SetTexture Stage, Texture

End Sub

Sub Set_EngineTextureEx(Texture As Direct3DTexture8, Optional ByVal Stage As Long = 0)

    obj_Device.SetTexture Stage, Texture

End Sub

Sub Set_EngineAmbientColor(ByVal color As Long)

    obj_Device.SetRenderState D3DRS_AMBIENT, color

    Data.Init_Ambient = color
    Data.Init_AmbientRGBA = LONGtoD3DCOLORVALUE(color)

End Sub

Sub Set_EngineTextureFilter(ByVal FILTER As QUEST3D_FILTERING_MODE, Optional ByVal ANISOTROPY_LEVEL As Long = 4)

  Dim Count As Integer

    For Count = 0 To Data.Capa.Max_TextureStages - 1

        If FILTER = QUEST3D_FILTER_POINT Then

            Call obj_Device.SetTextureStageState(Count, D3DTSS_MAGFILTER, D3DTEXF_POINT)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MINFILTER, D3DTEXF_POINT)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MIPFILTER, D3DTEXF_POINT)

            Data.TEXTURE_FILTER = D3DTEXF_POINT
        End If

        If FILTER = QUEST3D_FILTER_BILINEAR Then
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MAGFILTER, D3DTEXF_LINEAR)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MINFILTER, D3DTEXF_LINEAR)
            'Call obj_device.SetTextureStageState(count, D3DTSS_MIPFILTER, D3DTEXF_POINT)
            Data.TEXTURE_FILTER = D3DTEXF_LINEAR

        End If

        If FILTER = QUEST3D_FILTER_TRILINEAR And Data.Capa.Filter_Trilinear Then
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MAGFILTER, D3DTEXF_LINEAR)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MINFILTER, D3DTEXF_LINEAR)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MIPFILTER, D3DTEXF_LINEAR)

            Data.TEXTURE_FILTER = D3DTEXF_LINEAR

        End If

        If FILTER = QUEST3D_FILTER_GAUSSIANCUBIC And Data.Capa.Filetr_FlatCubic Then
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MAGFILTER, D3DTEXF_GAUSSIANCUBIC)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MINFILTER, D3DTEXF_LINEAR)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MIPFILTER, D3DTEXF_LINEAR)

            Data.TEXTURE_FILTER = D3DTEXF_GAUSSIANCUBIC

        End If

        If FILTER = QUEST3D_FILTER_FLATCUBIC And Data.Capa.Filter_GaussianCubic Then
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MAGFILTER, D3DTEXF_FLATCUBIC)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MINFILTER, D3DTEXF_LINEAR)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MIPFILTER, D3DTEXF_LINEAR)

            Data.TEXTURE_FILTER = D3DTEXF_FLATCUBIC
        End If

        If FILTER = QUEST3D_FILTER_ANISOTROPIC And Data.Capa.Filter_Anisotropic Then
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MIPFILTER, D3DTEXF_LINEAR)
            Call obj_Device.SetTextureStageState(Count, D3DTSS_MAXANISOTROPY, ANISOTROPY_LEVEL)

            Data.TEXTURE_FILTER = D3DTEXF_ANISOTROPIC
        End If

    Next Count

End Sub

Sub Pop_AllRenderstate()

  Dim I As Long

    I = 7
    While I <= 171

        If I <> 10 And I <> 11 And I <> 12 And I <> 13 And I <> 17 And I <> 18 And _
           I <> 21 And ((I < 31) Or (I > 39)) And ((I < 41) Or (I > 46)) And _
           ((I < 49) Or (I > 51)) And ((I < 61) Or (I > 127)) And I <> 144 And I <> 149 And I <> 150 And I <> 169 Then _
           obj_Device.SetRenderState I, Data.lpState.m_State(I)
        I = I + 1
    Wend

    obj_Device.SetMaterial Data.lpState.MATERIAL_state
    obj_Device.SetTransform D3DTS_PROJECTION, Data.lpState.PROJMAT_state
    obj_Device.SetTransform D3DTS_VIEW, Data.lpState.VIEWMAT_state
    obj_Device.SetTransform D3DTS_WORLD, Data.lpState.WORLDMAT_state

    Pop_AllTextureRenderstate

End Sub

Sub Pop_AllTextureRenderstate(Optional ByVal NumOfStage As Long = 2)

  Dim I As Integer
  Dim J As Integer

    For I = 0 To NumOfStage - 1
        For J = 1 To 28
            obj_Device.SetTextureStageState I, J, Data.lpState.T_state(I, J)
        Next J
    Next I

End Sub

Sub Pop_PixelShader()

    obj_Device.SetPixelShader Data.lpState.PixelShader

End Sub

'----------------------------------------
'Name: Pop_Renderstate
'Object: Pop
'Event: Renderstate
'----------------------------------------
'----------------------------------------
'Name: Pop_Renderstate
'Object: Pop
'Event: Renderstate
'Description:
'----------------------------------------
Sub Pop_Renderstate(ByVal StateVal As QUEST3D_RENDERSTATE)

    If StateVal = QUEST3DRS_MATERIAL Then
        obj_Device.SetMaterial Data.lpState.MATERIAL_state
        Exit Sub
    End If

    If StateVal = QUEST3DRS_MATRIX_PROJECTION Then
        obj_Device.SetTransform D3DTS_PROJECTION, Data.lpState.PROJMAT_state
        Exit Sub
    End If

    If StateVal = QUEST3DRS_MATRIX_VIEW Then
        obj_Device.SetTransform D3DTS_VIEW, Data.lpState.VIEWMAT_state
        Exit Sub
    End If

    If StateVal = QUEST3DRS_MATRIX_WORLD Then
        obj_Device.SetTransform D3DTS_WORLD, Data.lpState.WORLDMAT_state
        Exit Sub
    End If

    obj_Device.SetRenderState StateVal, Data.lpState.m_State(StateVal)

End Sub

Sub Pop_TextureRenderstate(ByVal Stage As Long, ByVal StateVal As QUEST3D_TEXTURERENDERSTATE)

    obj_Device.SetTextureStageState Stage, StateVal, Data.lpState.T_state(Stage, StateVal)

End Sub

Sub Pop_VertexShader()

    obj_Device.SetVertexShader Data.lpState.VertexShader

End Sub

Sub Push_AllRenderstate()

  Dim I As Long

    I = 7
    While I <= 171

        If I <> 10 And I <> 11 And I <> 12 And I <> 13 And I <> 17 And I <> 18 And _
           I <> 21 And ((I < 31) Or (I > 39)) And ((I < 41) Or (I > 46)) And _
           ((I < 49) Or (I > 51)) And ((I < 61) Or (I > 127)) And I <> 144 And I <> 149 And I <> 150 And I <> 169 Then _
           Data.lpState.m_State(I) = obj_Device.GetRenderState(I)
        I = I + 1
    Wend

    obj_Device.GetMaterial Data.lpState.MATERIAL_state
    obj_Device.GetTransform D3DTS_PROJECTION, Data.lpState.PROJMAT_state
    obj_Device.GetTransform D3DTS_VIEW, Data.lpState.VIEWMAT_state
    obj_Device.GetTransform D3DTS_WORLD, Data.lpState.WORLDMAT_state

    Push_AllTextureRenderstate

End Sub

Sub Push_AllTextureRenderstate(Optional ByVal NumOfStage As Long = 2)

  Dim I As Integer
  Dim J As Integer

    For I = 0 To NumOfStage - 1
        For J = 1 To 28
            Data.lpState.T_state(I, J) = obj_Device.GetTextureStageState(I, J)
        Next J
    Next I

End Sub

Sub Push_PixelShader()

    Data.lpState.PixelShader = obj_Device.GetPixelShader

End Sub

Sub Push_Renderstate(ByVal StateVal As QUEST3D_RENDERSTATE)

    If StateVal = QUEST3DRS_MATERIAL Then
        obj_Device.GetMaterial Data.lpState.MATERIAL_state
        Exit Sub
    End If

    If StateVal = QUEST3DRS_MATRIX_PROJECTION Then
        obj_Device.GetTransform D3DTS_PROJECTION, Data.lpState.PROJMAT_state
        Exit Sub
    End If

    If StateVal = QUEST3DRS_MATRIX_VIEW Then
        obj_Device.GetTransform D3DTS_VIEW, Data.lpState.VIEWMAT_state
        Exit Sub
    End If

    If StateVal = QUEST3DRS_MATRIX_WORLD Then
        obj_Device.GetTransform D3DTS_WORLD, Data.lpState.WORLDMAT_state
        Exit Sub
    End If

    Data.lpState.m_State(StateVal) = obj_Device.GetRenderState(StateVal)

End Sub

Sub Push_TextureRenderstate(ByVal Stage As Long, ByVal StateVal As QUEST3D_TEXTURERENDERSTATE)

    Data.lpState.T_state(Stage, StateVal) = obj_Device.GetTextureStageState(Stage, StateVal)

End Sub

'----------------------------------------
'Name: Set_EngineTextureFilter
'Object: Set
'Event: EngineTextureFilter
'Description:
'----------------------------------------

Sub Set_EngineTextureMipMapFilter(ByVal MimapFilter As QUEST3D_TEXTURE_FILTER)

    Data.TEXTURE_MIPMAPFILTER = MimapFilter

End Sub

'===============================================================
'
'Matrix state
'===============================================================
Sub Set_WorldRotate_Scale_TranslateEx(RotRadVec As D3DVECTOR, ScalVec As D3DVECTOR, TransVec As D3DVECTOR)

  Dim MatTemp As D3DMATRIX

    'D3DXMatrixIdentity MatTemp
    Call Matrix_Ret(MatTemp, ScalVec.x, ScalVec.y, ScalVec.z, RotRadVec.x, RotRadVec.y, RotRadVec.z, TransVec.x, TransVec.y, TransVec.z)
    obj_Device.SetTransform D3DTS_WORLD, MatTemp

End Sub

'----------------------------------------
'Name: Set_WorldRotate_Scale_TranslateEX
'Object: Set
'Event: WorldRotate_Scale_TranslateEX
'----------------------------------------
'----------------------------------------
'Name: Set_WorldRotate_Scale_TranslateEX
'Object: Set
'Event: WorldRotate_Scale_TranslateEX
'Description:
'----------------------------------------
Sub Set_WorldRotate_Scale_Translate(ByVal RadX As Single, ByVal RadY As Single, ByVal RadZ As Single, ByVal ScalX As Single, ByVal ScalY As Single, ByVal ScalZ As Single, ByVal TransX As Single, ByVal TransY As Single, ByVal TransZ As Single)

    Call Set_WorldRotate_Scale_TranslateEx(Vector(RadX, RadY, RadZ), Vector(ScalX, ScalY, ScalZ), Vector(TransX, TransY, TransZ))

End Sub

Sub Set_ViewFrustum(ByVal Near As Single, ByVal Far As Single, ByVal FovAngle As Single, Optional ByVal Aspect As Single = -1)

    If Aspect = -1 Then Aspect = Data.Buffer_Width / Data.Buffer_Height
    With Data.ViewFrust
        .Aspect = Aspect 'Data.Buffer_Height / Data.Buffer_Width
        .Near = Near
        .Far = Far
        .FovAngle = FovAngle

    End With

    D3DXMatrixPerspectiveFovLH Data.MatProjec, Data.ViewFrust.FovAngle, Data.ViewFrust.Aspect, Data.ViewFrust.Near, Data.ViewFrust.Far

    obj_Device.SetTransform D3DTS_PROJECTION, Data.MatProjec

End Sub

'================================

'----------------------------------------
'Name: Get_Capabilities
'Object: Get
'Event: Capabilities
'----------------------------------------
'----------------------------------------
'Name: Get_Capabilities
'Object: Get
'Event: Capabilities
'Description:
'----------------------------------------
Private Sub Get_Capabilities()

  Dim d3dCaps As D3DCAPS8, I As Integer, J As Integer

    obj_Device.GetDeviceCaps d3dCaps

    'check bump mapping

    ''//Does this device support the two bump mapping blend operations?
    If (d3dCaps.TextureOpCaps And D3DTEXOPCAPS_BUMPENVMAPLUMINANCE) Then
        Data.Capa.CanDo_BumpMapping = 1
    End If

    ''//Does this device support up to three blending stages?
    If d3dCaps.MaxTextureBlendStages < 3 Then
        Data.Capa.CandDo_3StagesTextureBlending = 0
      Else
        Data.Capa.CandDo_3StagesTextureBlending = 1

    End If

    ''//Does this device support multitexturing
    If d3dCaps.MaxSimultaneousTextures > 1 Then
        Data.Capa.CanDo_MultiTexture = 1
        Data.Capa.Max_TextureStages = d3dCaps.MaxSimultaneousTextures
    End If

    'anisotropic filter
    If d3dCaps.RasterCaps And D3DPRASTERCAPS_ANISOTROPY Then
        Data.Capa.Filter_Anisotropic = True

        Data.Capa.Max_AnisotropY = d3dCaps.MaxAnisotropy

    End If

    'trilinear

    If (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MINFLINEAR) Then

        If (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MAGFLINEAR) Then
            If (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MIPFLINEAR) Then

                Data.Capa.Filter_Trilinear = 1

            End If
        End If
    End If

    'flatcubic

    If ((d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MINFLINEAR) + _
       (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MAGFAFLATCUBIC) + _
       (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MIPFLINEAR)) Then

        Data.Capa.Filetr_FlatCubic = 1

    End If

    'Gaussian cubic

    If ((d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MINFLINEAR) + _
       (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MAGFGAUSSIANCUBIC) + _
       (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MIPFLINEAR)) Then

        Data.Capa.Filter_GaussianCubic = 1

    End If

    If d3dCaps.TextureCaps And D3DPTEXTURECAPS_VOLUMEMAP Then

        Data.Capa.CanDo_VolumeTexture = 1

    End If

    If d3dCaps.TextureCaps And D3DPTEXTURECAPS_PROJECTED Then

        Data.Capa.CanDo_ProjectedTexture = 1

    End If

    If d3dCaps.TextureCaps And D3DPTEXTURECAPS_MIPMAP Then

        Data.Capa.CanDo_TextureMipMapping = 1

    End If

    If (d3dCaps.RasterCaps And D3DPRASTERCAPS_WBUFFER) Then
        Data.Capa.Wbuffer_OK = True
        obj_Device.SetRenderState D3DRS_ZENABLE, D3DZB_USEW
        IS_WBUFFER = True
    End If

    If d3dCaps.MaxPointSize > 0 Then
        Data.Capa.CanDo_PointSprite = 1

    End If

  Dim MA As Long
  Dim MI As Long

    MA = D3DSHADER_VERSION_MAJOR(d3dCaps.VertexShaderVersion)
    MI = D3DSHADER_VERSION_MINOR(d3dCaps.VertexShaderVersion)

    'MA = D3DVS_VERSION(MA, MI)
    Data.Capa.Vertex_ShaderVERSION = Str(MI) + "." + CStr(MA)

    MA = D3DSHADER_VERSION_MAJOR(d3dCaps.PixelShaderVersion)
    MI = D3DSHADER_VERSION_MINOR(d3dCaps.PixelShaderVersion)

    'MA = D3DVS_VERSION(MA, MI)
    Data.Capa.Pixel_ShaderVERSIOn = Str(MI) + "." + CStr(MA)

    Data.Capa.Cando_VertexShader = d3dCaps.VertexShaderVersion >= D3DVS_VERSION(1, 0)
    Data.Capa.Cando_PixelShader = d3dCaps.PixelShaderVersion >= D3DVS_VERSION(1, 0)

    Data.Capa.CanDo_CubeMapping = (d3dCaps.TextureCaps And D3DPTEXTURECAPS_CUBEMAP)

    Data.Capa.CanDo_Dot3 = (d3dCaps.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3)

    Data.Capa.CanDoTableFog = (d3dCaps.RasterCaps And D3DPRASTERCAPS_FOGTABLE) And _
                              (D3DPRASTERCAPS_ZFOG) Or (d3dCaps.RasterCaps And D3DPRASTERCAPS_WFOG)

    Data.Capa.CanDoVertexFog = (d3dCaps.RasterCaps And D3DPRASTERCAPS_FOGVERTEX)

    Data.Capa.CanDoWFog = (d3dCaps.RasterCaps And D3DPRASTERCAPS_WFOG)

  Dim nAdapters As Long 'How many adapters we found
  Dim AdapterInfo As D3DADAPTER_IDENTIFIER8 'A Structure holding information on the adapter

  Dim sTemp As String

    '//This'll either be 1 or 2
    nAdapters = obj_D3D.GetAdapterCount

    For I = 0 To nAdapters - 1
        'Get the relevent Details
        obj_D3D.GetAdapterIdentifier I, 0, AdapterInfo

        'Get the name of the current adapter - it's stored as a long
        'list of character codes that we need to parse into a string
        ' - Dont ask me why they did it like this; seems silly really :)
        sTemp = "" 'Reset the string ready for our use

        For J = 0 To 511
            sTemp = sTemp & Chr$(AdapterInfo.Description(J))
        Next J
        sTemp = Replace(sTemp, Chr$(0), " ")
        J = InStr(sTemp, "     ")
        sTemp = Left$(sTemp, J)

    Next I

    If InStr(UCase(sTemp), "GEFORCE") Then

        If Data.Capa.Wbuffer_OK = 0 Then
            Data.Capa.Wbuffer_OK = 1
            IS_WBUFFER = 1
        End If

    End If

End Sub
