Attribute VB_Name = "Module_TextureManagement"
Option Explicit
'==============================================================================================================
'
'       In this module we define all objects and types
'       required for textures management
'
'
'
'This texture contains hardcoded routines for Texture loading
'
'
'
'
'
'
'
'
'
'================================================================================================
Private Type BITMAPHEADER
    'intMagic As Integer
    lngSize As Long
    intReserved1 As Integer
    intReserved2 As Integer
    lngOffset As Long

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private BMH As BITMAPHEADER

Public Const MIP_LEVELS = D3DX_DEFAULT

Enum QUEST3D_TEXTURE_TYPE
    OPAQUE = 0
    TRANSPARENT = 1
    DOT3 = 2

End Enum

Enum QUEST3D_TEXTURE_FILTER
    QUEST3D_TEXTURE_FILTER_POINT = 1
    QUEST3D_TEXTURE_FILTER_LINEAR = 2
    QUEST3D_TEXTURE_FILTER_TRIANGLE = 3
    QUEST3D_TEXTURE_FILTER_BOX = 4
    QUEST3D_TEXTURE_DEFAULT = -1
    QUEST3D_TEXTURE_FILTER_MIRROR_U = 65536  '(&H10000)
    QUEST3D_TEXTURE_FILTER_MIRROR_V = 131072 '(&H20000)
    QUEST3D_TEXTURE_FILTER_MIRROR = 196608   '(&H30000)
    QUEST3D_TEXTURE_FILTER_DITHER = 524288   '(&H80000)

End Enum

Public Type tTEXPOOL
    NumTextureInpool As Long
    NumLighmapsInpool As Long

    POOL_texture() As Direct3DTexture8
    POOL_Lightmaps() As Direct3DTexture8

    TextureName() As String

    Pool_TextureType() As QUEST3D_TEXTURE_TYPE

End Type

Public myTEXPOOL As tTEXPOOL

'====================================================================================
'
'
'
'===================================================================================
Function Add_TextureEx(Tfile As String, Optional ByVal Colorkey As Long = -1) As Long

    Add_TextureEx = Add_TextureToPoolEX2(CreateTextureColorKEY(Tfile, , , Colorkey))

End Function

Function Add_TextureFromMemory(BufferByte() As Byte, ByVal Wi As Integer, ByVal HI As Integer) As Long

    Add_TextureFromMemory = Add_TextureToPoolEX(CreateTextureFromBuffer(BufferByte, Wi, HI))

End Function

Function Add_TextureToPool(TexName As String, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal ColorKeyARGB As Long = -1) As Long

    Add_TextureToPool = -1

    Add_TextureToPool = Is_Inpool(TexName)

    If Add_TextureToPool > -1 Then

        Exit Function
    End If

    If InStr(UCase(TexName), ".WAL") > 0 Then
        '
        '        If FileiS_valid(TexName) Then
        '        myTEXPOOL.NumTextureInpool = myTEXPOOL.NumTextureInpool + 1
        '        ReDim Preserve myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1)
        '        ReDim Preserve myTEXPOOL.Pool_TextureType(myTEXPOOL.NumTextureInpool - 1)
        '        ReDim Preserve myTEXPOOL.TextureName(myTEXPOOL.NumTextureInpool - 1)
        '        End If
        '
        '        myTEXPOOL.TextureName(myTEXPOOL.NumTextureInpool - 1) = TexName
        '
        '        myTEXPOOL.Pool_TextureType(myTEXPOOL.NumTextureInpool - 1) = OPAQUE
        '
        '        Add_TextureToPool = Add_TextureToPoolEX(CreateTextureFromWAL(TexName, 0, 0))
        Exit Function

      ElseIf InStr(UCase(TexName), ".PCX") > 0 Then

        '        If FileiS_valid(TexName) Then
        '        myTEXPOOL.NumTextureInpool = myTEXPOOL.NumTextureInpool + 1
        '        ReDim Preserve myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1)
        '        ReDim Preserve myTEXPOOL.Pool_TextureType(myTEXPOOL.NumTextureInpool - 1)
        '        ReDim Preserve myTEXPOOL.TextureName(myTEXPOOL.NumTextureInpool - 1)
        '        End If
        '
        '        Add_TextureToPool = Add_TextureToPoolEX(Init.LoadPCX(TexName, Width, Height, ColorKeyARGB))
        Exit Function

    End If

    Add_TextureToPool = -1
    If FileIs_Valid(TexName) Then
        myTEXPOOL.NumTextureInpool = myTEXPOOL.NumTextureInpool + 1
        ReDim Preserve myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1)
        ReDim Preserve myTEXPOOL.Pool_TextureType(myTEXPOOL.NumTextureInpool - 1)
        ReDim Preserve myTEXPOOL.TextureName(myTEXPOOL.NumTextureInpool - 1)

        myTEXPOOL.TextureName(myTEXPOOL.NumTextureInpool - 1) = TexName

        myTEXPOOL.Pool_TextureType(myTEXPOOL.NumTextureInpool - 1) = OPAQUE

        If ColorKeyARGB <> 0 Then
            Set myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1) = CreateTextureColorKEY(TexName, Width, Height, ColorKeyARGB)
          Else
            Set myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1) = obj_D3DX.CreateTextureFromFileEx(obj_Device, TexName, 0, 0, MIP_LEVELS, 0, CFG.BK_FMT, D3DPOOL_MANAGED, Data.TEXTURE_FILTER, Data.TEXTURE_MIPMAPFILTER, 0, ByVal 0, ByVal 0)
        End If

        obj_D3DX.FilterTexture myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1), ByVal 0, 0, Data.TEXTURE_FILTER

        Add_TextureToPool = myTEXPOOL.NumTextureInpool - 1

    End If

End Function

Function Add_TextureToPoolEX(TEX As Direct3DBaseTexture8) As Long

  'If FileiS_valid(Tex) Then

    If TEX Is Nothing Then
        Add_TextureToPoolEX = -1
        Exit Function
    End If

    myTEXPOOL.NumTextureInpool = myTEXPOOL.NumTextureInpool + 1
    ReDim Preserve myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1)

    ReDim Preserve myTEXPOOL.Pool_TextureType(myTEXPOOL.NumTextureInpool - 1)

    myTEXPOOL.Pool_TextureType(myTEXPOOL.NumTextureInpool - 1) = OPAQUE

    Set myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1) = TEX
    obj_D3DX.FilterTexture myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1), ByVal 0, 0, Data.TEXTURE_FILTER

  Dim dd As D3DSURFACE_DESC
    Call myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1).GetLevelDesc(0, dd)
    
    
    Add_TextureToPoolEX = myTEXPOOL.NumTextureInpool - 1
    'End If

End Function

Function Add_TextureToPoolEX2(TEX As Direct3DTexture8) As Long

  'If FileiS_valid(Tex) Then

    myTEXPOOL.NumTextureInpool = myTEXPOOL.NumTextureInpool + 1
    ReDim Preserve myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1)

    ReDim Preserve myTEXPOOL.Pool_TextureType(myTEXPOOL.NumTextureInpool - 1)

    myTEXPOOL.Pool_TextureType(myTEXPOOL.NumTextureInpool - 1) = OPAQUE

    Set myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1) = TEX
    obj_D3DX.FilterTexture myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1), ByVal 0, 0, Data.TEXTURE_FILTER

  Dim dd As D3DSURFACE_DESC
    Call myTEXPOOL.POOL_texture(myTEXPOOL.NumTextureInpool - 1).GetLevelDesc(0, dd)
    
    

    Add_TextureToPoolEX2 = myTEXPOOL.NumTextureInpool - 1
    'End If

End Function

Function Is_Inpool(fl As String) As Integer

  Dim I As Long

    Is_Inpool = -1

    If myTEXPOOL.NumTextureInpool = 0 Then Exit Function

    On Error GoTo Err
    For I = 0 To UBound(myTEXPOOL.TextureName)

        If fl = myTEXPOOL.TextureName(I) Then
            Is_Inpool = I
            Exit Function
        End If
    Next I

Exit Function

Err:

End Function

Public Function CreateTextureColorKEY(ByVal TextureFile As String, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal ColorKeyARGB As Long = -1) As Direct3DTexture8

  'On Error Resume Next

  'Set tex = obj_d3dx.CreateTextureFromFileEx(obj_device, Filename, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, ColorKeyARGB, ByVal 0, ByVal 0)
  'Debug.Print Filename

  Dim TEX As Direct3DTexture8
  Dim Wi As Integer, HI As Integer

    If FileIs_Valid(TextureFile) = False Then Exit Function

    Set TEX = obj_D3DX.CreateTextureFromFileEx(obj_Device, _
        TextureFile, Width, Height, 1, 0, D3DFMT_X8R8G8B8, D3DPOOL_MANAGED, _
        Data.TEXTURE_MIPMAPFILTER, Data.TEXTURE_MIPMAPFILTER, ColorKeyARGB, ByVal 0, _
        ByVal 0)

  Dim DES As D3DSURFACE_DESC
    Call TEX.GetLevelDesc(0, DES)
  Dim pData As D3DLOCKED_RECT

    If Width = -1 Then Width = DES.Width
    If Height = -1 Then Height = DES.Height

    '    WI = width
    '    HI = Height

  Dim I, J
    I = DES.Width
  Dim ARR() As Byte

  Dim A, R, g, b
    '
  Dim x As Long, y As Long, px As Long
  Dim pxArr() As Byte, px1 As Byte, px2 As Byte
  Dim lRes As Long, bFirst As Byte, bSecond As Byte, lFirst As Long, lSecond As Long
  Dim bRed As Byte, bGreen As Byte, bBlue As Byte
  Dim lRed As Long, lGreen As Long, lBlue As Long

    TEX.LockRect 0, pData, ByVal 0, 0

    'ReDim ARR(pData.Pitch * DES.Height)
    '
    'we can now play around with the stuff in pData
    ReDim pxArr(4) As Byte 'enough bytes
    If Not (DXCopyMemory(pxArr(0), ByVal pData.pBits, 4) = D3D_OK) Then Debug.Print "COPY TO ARRAY FAILED"

    'Should be XRGB format, instead, in BGRX format...
    For x = 0 To 3 Step 4
        'unused = pxArr(x + 3)
        bRed = pxArr(x + 2)
        bGreen = pxArr(x + 1)
        bBlue = pxArr(x + 0)
        'I dont want to change the colours.... :)
        'bRed = 0
        'bGreen = 0
        'bBlue = 255
        pxArr(x + 2) = bRed
        pxArr(x + 1) = bGreen
        pxArr(x + 0) = bBlue
    Next x

    TEX.UnlockRect 0

    If ColorKeyARGB = -1 Then

        'A = 255
        'r = 0
        'g = 0
        'b = 0
        ColorKeyARGB = D3DColorXRGB(bRed, bGreen, bBlue)
      Else

    End If

    '    Dim FMT As Long
    '    Call TEX.GetLevelDesc(0, DES)
    '    FMT = DES.format

    Set TEX = obj_D3DX.CreateTextureFromFileEx(obj_Device, _
        TextureFile, Width, Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKeyARGB, ByVal 0, ByVal 0)

    Set CreateTextureColorKEY = TEX
    Set TEX = Nothing

End Function

'----------------------------------------
'Name: CreateTextureFromBuffer
'----------------------------------------
'----------------------------------------
'Name: CreateTextureFromBuffer
'Description:
'----------------------------------------
Function CreateTextureFromBuffer(BufferByte() As Byte, Wi, HI, Optional ByVal Colorkey As Long = &H0) As Direct3DBaseTexture8

  Dim Surf As Direct3DBaseTexture8
  Dim MAP() As Byte
  Dim I As Long
  Dim J As Long

    I = UBound(BufferByte)
    J = LBound(BufferByte)

    If I < 9 Then Exit Function

    ReDim MAP(CLng(Wi * HI * 3) + 54)

    With BMH
        '.intMagic = 19778
        .lngSize = (Wi * HI * 3) + 54
        .lngOffset = 54

        .biBitCount = 24
        .biWidth = Wi
        .biHeight = HI
        .biSize = 40
        .biPlanes = 1
        .biSizeImage = Wi * HI * 3
        .biCompression = 0
        .biXPelsPerMeter = 50
        .biYPelsPerMeter = 50
    End With

  Dim BMPheader(Len(BMH) + 1) As Byte
    CopyMemory BMPheader(2), BMH, Len(BMH)
    BMPheader(0) = 66
    BMPheader(1) = 77

    'Copy  data into an array
    CopyMemory MAP(0), BMPheader(0), 54
    CopyMemory MAP(54), BufferByte(J), Wi * HI * 3

    If Colorkey = &H0 Then
        Set Surf = obj_D3DX.CreateTextureFromFileInMemoryEx(obj_Device, MAP(0), UBound(MAP()), Wi, HI, MIP_LEVELS, 0, CFG.BK_FMT, D3DPOOL_DEFAULT, Data.TEXTURE_FILTER, Data.TEXTURE_MIPMAPFILTER, &H0, ByVal 0, ByVal 0)

      Else
        Set Surf = obj_D3DX.CreateTextureFromFileInMemoryEx(obj_Device, MAP(0), UBound(MAP()), Wi, HI, MIP_LEVELS, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, Data.TEXTURE_FILTER, Data.TEXTURE_MIPMAPFILTER, Colorkey, ByVal 0, ByVal 0)
    End If

    obj_D3DX.FilterTexture Surf, ByVal 0, 0, Data.TEXTURE_FILTER

    Set CreateTextureFromBuffer = Surf

End Function
