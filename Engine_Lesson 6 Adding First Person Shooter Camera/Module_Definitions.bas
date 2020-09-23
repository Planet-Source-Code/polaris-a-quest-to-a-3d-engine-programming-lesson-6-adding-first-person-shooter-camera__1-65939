Attribute VB_Name = "Module_Definitions"
Option Explicit
'==============================================================================================================
'
'       In this module we define all objects and types
'       we will use in the Engine like,3D Device object,iformations
'       like matricies ect....
'
'
'================================================================================================

'global objects

Public Const QUEST3D_PI As Double = 3.14159265358979

Public Const QUEST3D_RAD = QUEST3D_PI / 180

Public Const QUEST3D_EPSILON = 0.0005

Enum QUEST3D_RENDERSTATE
    QUEST3DRS_ZENABLE = 7
    QUEST3DRS_FILLMODE = 8
    QUEST3DRS_SHADEMODE = 9
    QUEST3DRS_LINEPATTERN = 10
    QUEST3DRS_ZWRITEENABLE = 14
    QUEST3DRS_ALPHATESTENABLE = 15
    QUEST3DRS_LASTPIXEL = 16
    QUEST3DRS_SRCBLEND = 19
    QUEST3DRS_DESTBLEND = 20
    QUEST3DRS_CULLMODE = 22
    QUEST3DRS_ZFUNC = 23
    QUEST3DRS_ALPHAREF = 24
    QUEST3DRS_ALPHAFUNC = 25
    QUEST3DRS_DITHERENABLE = 26
    QUEST3DRS_ALPHABLENDENABLE = 27
    QUEST3DRS_FOGENABLE = 28
    QUEST3DRS_SPECULARENABLE = 29
    QUEST3DRS_ZVISIBLE = 30
    QUEST3DRS_FOGCOLOR = 34
    QUEST3DRS_FOGTABLEMODE = 35
    QUEST3DRS_FOGSTART = 36
    QUEST3DRS_FOGEND = 37
    QUEST3DRS_FOGDENSITY = 38
    QUEST3DRS_EDGEANTIALIAS = 40
    QUEST3DRS_ZBIAS = 47
    QUEST3DRS_RANGEFOGENABLE = 48
    QUEST3DRS_STENCILENABLE = 52
    QUEST3DRS_STENCILFAIL = 53
    QUEST3DRS_STENCILZFAIL = 54
    QUEST3DRS_STENCILPASS = 55
    QUEST3DRS_STENCILFUNC = 56
    QUEST3DRS_STENCILREF = 57
    QUEST3DRS_STENCILMASK = 58
    QUEST3DRS_STENCILWRITEMASK = 59
    QUEST3DRS_TEXTUREFACTOR = 60
    QUEST3DRS_WRAP0 = 128
    QUEST3DRS_WRAP1 = 129
    QUEST3DRS_WRAP2 = 130
    QUEST3DRS_WRAP3 = 131
    QUEST3DRS_WRAP4 = 132
    QUEST3DRS_WRAP5 = 133
    QUEST3DRS_WRAP6 = 134
    QUEST3DRS_WRAP7 = 135
    QUEST3DRS_CLIPPING = 136
    QUEST3DRS_LIGHTING = 137
    QUEST3DRS_AMBIENT = 139
    QUEST3DRS_FOGVERTEXMODE = 140
    QUEST3DRS_COLORVERTEX = 141
    QUEST3DRS_LOCALVIEWER = 142
    QUEST3DRS_NORMALIZENORMALS = 143
    QUEST3DRS_DIFFUSEMATERIALSOURCE = 145
    QUEST3DRS_SPECULARMATERIALSOURCE = 146
    QUEST3DRS_AMBIENTMATERIALSOURCE = 147
    QUEST3DRS_EMISSIVEMATERIALSOURCE = 148
    QUEST3DRS_VERTEXBLEND = 151
    QUEST3DRS_CLIPPLANEENABLE = 152
    QUEST3DRS_SOFTWAREVERTEXPROCESSING = 153
    QUEST3DRS_POINTSIZE = 154
    QUEST3DRS_POINTSIZE_MIN = 155
    QUEST3DRS_POINTSPRITEENABLE = 156
    QUEST3DRS_POINTSCALEENABLE = 157
    QUEST3DRS_POINTSCALE_A = 158
    QUEST3DRS_POINTSCALE_B = 159
    QUEST3DRS_POINTSCALE_C = 160
    QUEST3DRS_MULTISAMPLEANTIALIAS = 161
    QUEST3DRS_MULTISAMPLEMASK = 162
    QUEST3DRS_PATCHEDGESTYLE = 163
    QUEST3DRS_PATCHSEGMENTS = 164
    QUEST3DRS_DEBUGMONITORTOKEN = 165
    QUEST3DRS_POINTSIZE_MAX = 166
    QUEST3DRS_INDEXEDVERTEXBLENDENABLE = 167
    QUEST3DRS_COLORWRITEENABLE = 168
    QUEST3DRS_TWEENFACTOR = 170
    QUEST3DRS_BLENDOP = 171
    QUEST3DRS_POSITIONORDER = 172
    QUEST3DRS_NORMALORDER = 173
    QUEST3DRS_MATRIX_WORLD = 174
    QUEST3DRS_MATRIX_VIEW = 175
    QUEST3DRS_MATRIX_PROJECTION = 176
    QUEST3DRS_MATERIAL = 177

End Enum

Enum QUEST3D_TEXTURERENDERSTATE

    QUEST3DTSS_COLOROP = 1 '
    QUEST3DTSS_COLORARG1 = 2 '
    QUEST3DTSS_COLORARG2 = 3 '
    QUEST3DTSS_ALPHAOP = 4 '
    QUEST3DTSS_ALPHAARG1 = 5 '
    QUEST3DTSS_ALPHAARG2 = 6 '
    QUEST3DTSS_BUMPENVMAT00 = 7 '
    QUEST3DTSS_BUMPENVMAT01 = 8 '
    QUEST3DTSS_BUMPENVMAT10 = 9 '
    QUEST3DTSS_BUMPENVMAT11 = 10 '
    QUEST3DTSS_TEXCOORDINDEX = 11 '
    QUEST3DTSS_ADDRESSU = 13 '
    QUEST3DTSS_ADDRESSV = 14 '
    QUEST3DTSS_BORDERCOLOR = 15 '
    QUEST3DTSS_MAGFILTER = 16             '   //(0x10)
    QUEST3DTSS_MINFILTER = 17             '   //(0x11)
    QUEST3DTSS_MIPFILTER = 18             '   //(0x12)
    QUEST3DTSS_MIPMAPLODBIAS = 19         '   //(0x13)
    QUEST3DTSS_MAXMIPLEVEL = 20           '   //(0x14)
    QUEST3DTSS_MAXANISOTROPY = 21         '   //(0x15)
    QUEST3DTSS_BUMPENVLSCALE = 22         '   //(0x16)
    QUEST3DTSS_BUMPENVLOFFSET = 23        '   //(0x17)
    QUEST3DTSS_TEXTURETRANSFORMFLAGS = 24 '   //(0x18)
    QUEST3DTSS_ADDRESSW = 25              '   //(0x19)
    QUEST3DTSS_COLORARG0 = 26             '   //(0x1A)
    QUEST3DTSS_ALPHAARG0 = 27             '   //(0x1B)
    QUEST3DTSS_RESULTARG = 28             '(&H1C)

End Enum

'for accessing to all functions provided by Directx Lib
Public obj_DX As New DirectX8

'this object is an interface that provide functions and methods.
'this routines allow to check if the real 3D device has some required
'capabilities for a 3D engine
Public obj_D3D As Direct3D8

'this engine is an interface that communicate
'directly with the 3D GFX Card
Public obj_Device As Direct3DDevice8

Public obj_D3DX As D3DX8

'=======================================================================
' here we define all type that will be required
'
'
'=======================================================================

Public Type QUEST3D_FOV
    Near As Single
    Far As Single
    FovAngle As Single
    Aspect As Single

End Type

Type QUEST3D_SaveState
    m_State(7 To 171) As Long
    Init_Renderstate(7 To 171) As Long

    T_state(0 To 7, 1 To 28) As Long
    Init_TexSate(0 To 7, 1 To 28) As Long

    MATERIAL_state As D3DMATERIAL8
    VIEWMAT_state As D3DMATRIX
    PROJMAT_state As D3DMATRIX
    WORLDMAT_state As D3DMATRIX

    PixelShader As Long
    VertexShader As Long

End Type

Public Type QUEST3D_CAPABILITIES
    Filter_Bilinear As Boolean
    Filter_Trilinear As Boolean
    Filter_Anisotropic As Boolean
    Filter_GaussianCubic As Boolean
    Filetr_FlatCubic As Boolean

    CanDo_MultiTexture As Boolean
    CanDo_CubeMapping As Boolean
    CanDo_Dot3 As Boolean
    CanDo_VolumeTexture As Boolean
    CanDo_ProjectedTexture As Boolean
    CanDo_TextureMipMapping As Boolean
    CanDo_PureDevice As Boolean
    CanDo_PointSprite As Boolean

    Cando_RenderSurface As Boolean
    CandDo_3StagesTextureBlending As Boolean

    Cando_PixelShader As Boolean
    Cando_VertexShader As Boolean

    CanDoTableFog        As Boolean
    CanDoVertexFog       As Boolean
    CanDoWFog            As Boolean

    TandL_Device As Boolean
    CanDo_BumpMapping As Boolean

    Wbuffer_OK As Boolean
    Max_ActiveLights As Long
    Max_TextureStages As Long
    Max_AnisotropY As Long

    Pixel_ShaderVERSIOn As String

    Vertex_ShaderVERSION As String

End Type


'for View Frustum
Public Type QUEST3D_FRUSTUM
   PLANE(5) As D3DPLANE
End Type

Public Enum QUEST3D_FRUSTUMSIDE

    QUEST3D_RIGHT = 0          '        // The RIGHT side of the frustum
    QUEST3D_LEFT = 1           '        // The LEFT  side of the frustum
    QUEST3D_BOTTOM = 2         '        // The BOTTOM side of the frustum
    QUEST3D_TOP = 3                '        // The TOP side of the frustum
    QUEST3D_BACK = 4           '        // The BACK side of the frustum
    QUEST3D_FRONT = 5          '// The FRONT side of the frustum
End Enum


Public lpFRUST As QUEST3D_FRUSTUM


 Type CAM
    EYE As D3DVECTOR
    RotVec As D3DVECTOR
    DEG As Single
    AngX As Single
    ANGy As Single
    ANGz As Single
    Dest_at As D3DVECTOR
    Dir As D3DVECTOR
    DirNormalized As D3DVECTOR

    Mouse_Prev_X As Long
    Mouse_Prev_Y As Long

End Type

Public Type QUEST3D_CFG

    'actual  screen width
    Buffer_Width As Integer
    'actual screen_height
    Buffer_Height As Integer
    'screen Rectangle (left,right,top,bottom values)
    Buffer_Rect As RECT
    'are we in windowed mode
    Is_Windowed As Boolean
    'dephtbit size
    Bpp As Integer
    'the engine is active
    Is_engineActive As Boolean
    'color for the back buffer
    BackBuff_ClearColor As Long

    GamaLevel As Single

    'for font
    MainFont As D3DXFont
    StFont As StdFont
    FontDesc As IFont

    

    'handle of the form or the windows interface
    Hwindow As Long
    HwindowParent As Long

    'device creation parameters
    WinParam As D3DPRESENT_PARAMETERS

    'for frame counter
    Fps_CurrentTime As Single
    Fps_LastTime As Single
    Fps_FrameCounter As Single
    Fps_FramePerSecond As Single
    Fps_TimePassed As Single

    'for texture
    TEXTURE_FILTER As CONST_D3DTEXTUREFILTERTYPE
    TEXTURE_MIPMAPFILTER As QUEST3D_TEXTURE_FILTER

    'this is to avoid Backbuffer clearing
    'it can increase PFS
    IS_ClearRenderTarget As Boolean

    'this is for renderstate saving and recalling
    lpState As QUEST3D_SaveState

    'for polygons drawing stats
    Total_TriangleRENDERED As Long
    Total_VerticeRENDERED As Long

    'for material color
    Init_Ambient As Long
    Init_AmbientRGBA As D3DCOLORVALUE
    Init_Material As D3DMATERIAL8

    'for device capabilities

    Capa As QUEST3D_CAPABILITIES
    
    'for camera
    EYES As CAM
    
    'for view frustum
    ViewFrust As QUEST3D_FOV
    MatProjec As D3DMATRIX
    matView As D3DMATRIX
    FRUSTUM_HASCHANGED As Boolean
    
    'for input
    IS_DinputOK As Boolean
    IS_Joystick As Boolean 'are joysticks connected
    JoyNumDevice As Long 'number of joysticks
      'direct inpu properties for joystick
    DiProp_Dead As DIPROPLONG
    DiProp_Range As DIPROPRANGE
    DiProp_Saturation As DIPROPLONG

End Type

Public Data As QUEST3D_CFG

Public LpGLOBAL_QUEST3D As cQuest3D_Core

'this is added for
'dialog based initialization

'this will received FMT format
Type tDFMT
    NumD As Long
    FMT() As Long
End Type

Type tFMT
    Wi As Long
    HI As Long
    'FMT As Long
    BK_FMT() As Long
    DP_FMT() As tDFMT

    SAMPLE() As Long
    NumBK As Long
    NumDPH As Long
    NumD() As Long

End Type

Type Tini
    RESO() As tFMT
    SelectINDEX As Long
    NumRES As Long
    CurrentWINDOWED As tFMT
    DISP_FMT() As Long
    CurrBKFMT As Long
    CurrINDEX As Long

End Type

Public TempDX8 As DirectX8          'The Root Object
Public TempD3D8 As Direct3D8      'The Direct3D Interface

Public nAdapters As Long 'How many adapters we found
Public AdapterInfo As D3DADAPTER_IDENTIFIER8 'A Structure holding information on the adapter

Public nModes As Long 'How many display modes we found
Public DEV As Tini
Public Type QUEST3D_CFG_INI

    Width As Integer
    Height As Integer
    format As Long
    USE_from_Dialog As Boolean
    MaxFramePerSec As Long
    USE_TnL As Boolean
    DeviceTyp As CONST_D3DDEVTYPE
    ForceVerSINC As Boolean
    appHandle As Long
    ChildHandle As Long
    IS_FullScreen As Boolean
    GamaLevel As Single
    Bpp As Integer
    BufferCount As Integer
    BK_FMT As Long
    DP_FMT As Long
    IS_OKAY As Long

End Type

'check if there is ERROR
Public IS_ERROR As Boolean

Public IS_WBUFFER As Boolean

Public CFG As QUEST3D_CFG_INI

'some apis
'to retrieve keyboard state
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'to retrieve time
Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'some types

'for rendering

' A structure for a simple Transformed and lighted vertex type
'
'Transformed and lighet means
'the vertice is a point on the screen
'
'===================================================
'                       |Y-
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'X-_____________________x_________________________X+
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |Y+

'
'
'
' representing a point on the screen

Public Type QUEST3D_VERTEXCOLORED2D
    Position As D3DVECTOR
    'where
    '    x As Single         'x in screen space
    '    y As Single         'y in screen space
    '    z  As Single        'normalized z
    rhw As Single       'normalized z rhw
    color As Long       'vertex color
End Type

' Our custom FVF, which describes our custom vertex structure
Public Const QUEST3D_FVFVERTEXCOLORED2D = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE)

Public Type QUEST3D_VERTEXCOLORED3D
    Position As D3DVECTOR
    color As Long       'vertex color
End Type

' Our custom FVF, which describes our custom vertex structure
Public Const QUEST3D_FVFVERTEXCOLORED3D = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

'here is a new type for an Untransformed and unlighted Vertex
Public Type QUEST3D_VERTEX
    Position As D3DVECTOR
    Normal As D3DVECTOR
    Texture As D3DVECTOR2
End Type

' Our custom FVF, which describes our custom vertex structure
Public Const QUEST3D_FVFVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

'here is a new type for an Untransformed and unlighted Vertex
'with two texture stage
Public Type QUEST3D_VERTEX2
    Position As D3DVECTOR
    Normal As D3DVECTOR
    Texture1 As D3DVECTOR2
    Texture2 As D3DVECTOR2

End Type

' Our custom FVF, which describes our custom vertex structure
Public Const QUEST3D_FVFVERTEX2 = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX2)

'We will used this for matrix calculus
Public Type VerTEX_PARAM
    vPosition As D3DVECTOR
    Vscal As D3DVECTOR
    Vrotate As D3DVECTOR

    WorldMatrix As D3DMATRIX
    WorldInvMatrix As D3DMATRIX
    HasChanged As Boolean
    BoxIsComputed As Boolean

    ID As Long
End Type
