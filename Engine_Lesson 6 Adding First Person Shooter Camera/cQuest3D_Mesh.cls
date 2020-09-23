VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuest3D_Mesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'
'       THIS IS THE 3D MESH CLASS
'
'  - we start generate basic and complex 3D forms
'  - We add texture
'  - we Prepare Mesh datas for rendering
'  - we render Mesh
'  - we close when we are done
'================================================================================================

Option Explicit

'what we do
'we define all types required here
'
'3D OBJCET
'    SUB_OBJECT0
'        Polygons() As QUEST3D_VERTEX2
'        NumberOfpolygons
'        VertexBuffer for rendering the polys
'        checki if it's a hi poly object
'    SUB_OBJECT1
'        Polygons() As QUEST3D_VERTEX2
'        NumberOfpolygons
'        VertexBuffer for rendering the polys
'        checki if it's a hi poly object
'    SUB_OBJECT2
'        Polygons() As QUEST3D_VERTEX2
'        NumberOfpolygons
'        VertexBuffer for rendering the polys
'        checki if it's a hi poly object
'    SUB_OBJECT...normal.
'        Polygons() As QUEST3D_VERTEX2
'        NumberOfpolygons
'        VertexBuffer for rendering the polys
'        checki if it's a hi poly object

'the 3d object is divised by textures

'information for a sub mesh
'for rendering and collision
Private Type tSinglePOLY

    NumVert As Long
    Vert() As QUEST3D_VERTEX2
    ZZindex As Long
    OctreeID As Long
    TotalNodes As Long

    Total_VertVis As Long
    TotalVERT As Long
    TextureID As Long

    DXMESH As D3DXMesh

    DXMESH2 As D3DXMesh

    DXMESH3 As D3DXMesh

    DXMESH4 As D3DXMesh

    DXMESH5 As D3DXMesh
    
    DXMESH6 As D3DXMesh

    DXMESH7 As D3DXMesh

    DXMESH8 As D3DXMesh

    DXMESH9 As D3DXMesh
    
    DXMESH10 As D3DXMesh

    Num_Mesh As Byte

    'we define 10 chunks to
    'allow hi polygon scene management
    DXMESH_FAIL As Boolean

    DXMESH_FAIL2 As Boolean
    DXMESH_FAIL3 As Boolean
    DXMESH_FAIL4 As Boolean
    DXMESH_FAIL5 As Boolean
    DXMESH_FAIL6 As Boolean
    DXMESH_FAIL7 As Boolean
    DXMESH_FAIL8 As Boolean
    DXMESH_FAIL9 As Boolean
    
    

    VBUF As Direct3DVertexBuffer8
    iBuf As Direct3DIndexBuffer8
    nIndici As Long
    nFacce As Long

    vBuf2 As Direct3DVertexBuffer8
    iBuf2 As Direct3DIndexBuffer8
    nIndici2 As Long
    nFacce2 As Long

    vBuf3 As Direct3DVertexBuffer8
    iBuf3 As Direct3DIndexBuffer8
    nIndici3 As Long
    nFacce3 As Long

    SizeofVert As Long
End Type

Private Type QUEST3DMeshTextureInfo
    Filename As String * 255
End Type

Public Enum QUEST3DMeshTypePrimitive
    QUEST3D_SPHERE = &H10
    QUEST3D_TEAPOT = &H110
    QUEST3D_TAURUS = &H1110

End Enum

Private Type QUEST3DSubMesh
    Pos As D3DVECTOR
    Mesh As D3DXMesh
    MeshType As Long
    TEXID As Long
End Type

Private Type QUEST3DSubMeshSave
    Pos As D3DVECTOR
    TEXID As Long
    PARAM(10) As Single
    MeshType As Long
End Type

Public Event Progress(ProcessTitle As String, PercentDone As Integer)

Private Type QUEST3DMeshHeader
    ID As String * 4
    version As Long

    NumVert As Long
    NumText As Long
    numlightmaps As Long
    NumPOLY As Long
    NumPrimitives As Long
    NumSubMesh As Long
    numLight As Long
    POLY_unsorted As Long

    RSCullMode As CONST_D3DCULL
    IS_Builded As Boolean
    IS_LOADED As Boolean
    IS_OCTREE As Boolean

End Type

Private Type Tmesh
    Header As QUEST3DMeshHeader

    TextureInfo() As QUEST3DMeshTextureInfo
    TexturesID() As Long 'local texture index
    TexturesDOT3ID() As Long 'local texture index

    TexturesType() As Long 'text type
    LightmapsID() As Long 'local lighmap index
    PolyTexIndex() As Long

    POLY_sorted() As tSinglePOLY

    Num_Box As Long

    Vertex() As QUEST3D_VERTEX2

    VB() As Direct3DVertexBuffer8

    USE_DOT3 As Boolean

    Xdata As VerTEX_PARAM

    FVF As Long
    NumP As Long
    NumV As Long

End Type

Private Type TQUEST3DMesh
    ID As Long
    NumTextures As Long
    numLight As Long
    TextureInfo() As QUEST3DMeshTextureInfo
    TexturesID() As Long 'local texture index
    TexturesType() As Long 'text type
    POLY_sorted() As tSinglePOLY
End Type

Private Type tSinglePOLYSave

    NumVert As Long
    Vert() As QUEST3D_VERTEX2
End Type

Private Type TmeshSave
    Num_TextTures As Long
    Text() As String
    Polyz() As tSinglePOLYSave

End Type

Private OCTREE_VERT() As QUEST3D_VERTEX2
Private OCTREE_MAXVERT As Long

Private SubMesh() As QUEST3DSubMesh
Private NumSubMesh As Long
Private IS_ThereMesh As Boolean

Private OCTREE_MINPOLYZ As Long

Public Enum QUEST3DMESH_STAIRS_ORIENTATION
    SOUTH_NORTH = &H0
    NORTH_SOUTH = &H10
    EAST_WEST = &H110
    WEST_EAST = &H1110

End Enum

Private mMESH As Tmesh
Dim PolyCount As Long
Dim LoadFromFile As Boolean
Private TexturePATH As String

Sub Add_Vertex(x As Single, y As Single, z As Single, Optional ByVal nx As Single, Optional ByVal ny As Single, Optional ByVal nz As Single, Optional ByVal Tu As Single, Optional ByVal TV As Single, Optional ByVal TextureID As Long = 0, Optional ByVal LightmapID As Long = -1)

  Dim V As QUEST3D_VERTEX2

    V.Position.x = x
    V.Position.y = y
    V.Position.z = z
    V.Normal.x = nx
    V.Normal.y = ny
    V.Normal.z = nz
    V.Texture1.x = Tu
    V.Texture1.y = TV
    V.Texture2.x = Tu
    V.Texture2.y = TV

    With mMESH
        'On Error GoTo Err
DEB:
        If .Header.NumText - 1 < TextureID Then
            ReDim Preserve .POLY_sorted(TextureID)

            ReDim Preserve .TexturesID(TextureID)
            ReDim Preserve .TexturesType(TextureID)

            .Header.NumText = TextureID + 1
            .TexturesID(TextureID) = -1

        End If

        .POLY_sorted(TextureID).TotalVERT = .POLY_sorted(TextureID).TotalVERT + 1

        ReDim Preserve .POLY_sorted(TextureID).Vert(.POLY_sorted(TextureID).TotalVERT - 1)

        .POLY_sorted(TextureID).NumVert = .POLY_sorted(TextureID).TotalVERT

        mMESH.POLY_sorted(TextureID).Vert(.POLY_sorted(TextureID).TotalVERT - 1) = V

        If .POLY_sorted(TextureID).TotalVERT Mod 3 = 0 Then
            '          TempPOL.Vertex(0) = mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 3)
            '          TempPOL.Vertex(1) = mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 2)
            '          TempPOL.Vertex(2) = mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 1)
            '
            '          N = NemoPOLYGON_GetPolyNormal(TempPOL)
            '
            '          mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 3).Normal.x = N.Position.x
            '          mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 3).Normal.y = N.Position.y
            '          mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 3).Normal.z = N.Position.z
            '
            '           mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 2).Normal.x = N.Position.x
            '          mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 2).Normal.y = N.Position.y
            '          mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 2).Normal.z = N.Position.z
            '
            '           mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 1).Normal.x = N.Position.x
            '          mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 1).Normal.y = N.Position.y
            '          mMESH.POLY_sorted(TextureID).VERT(.POLY_sorted(TextureID).TotalVERT - 1).Normal.z = N.Position.z

        End If

        .NumV = .NumV + 1

    End With

Exit Sub

Err:

End Sub

'----------------------------------------
'Name: Add_VertexEX
'Object: Add
'Event: VertexEX
'----------------------------------------
Sub Add_VertexEx(vPos As D3DVECTOR, vNormal As D3DVECTOR, TexUV As D3DVECTOR2, Optional ByVal TextureID As Long = 0)

    Call Add_Vertex(vPos.x, vPos.y, vPos.z, vNormal.x, vNormal.y, vNormal.z, TexUV.x, TexUV.y, TextureID)

End Sub

'Sub Add_vertexColored(ByVal ArrayPointer As Long, ByVal NumVert As Long, Optional TextureID As Long = 0)
'
'  Dim Vert() As QUEST3D_VERTEXCOLORED3D
'
'    ReDim Vert(NumVert - 1)
'  Dim I As Long
'
'    CopyMemory Vert(0), ArrayPointer, Len(Vert(0)) * NumVert
'
'    For I = 0 To NumVert - 1
'        Me.Add_Vertex Vert(I).Position.x, Vert(I).Position.y, Vert(I).Position.z, , , , , , TextureID
'
'    Next I
'
'End Sub

Sub Add_Box(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional ByVal TextureIDLeft As Long = 0, Optional ByVal TextureIDRight As Long = 0, Optional ByVal TextureIDTop As Long = 0, Optional ByVal TextureIDBottom As Long = 0, Optional ByVal TextureIDFront As Long = 0, Optional ByVal TextureIDBack As Long = 0, Optional ByVal Tu As Single = 1, Optional ByVal TV As Single = 1)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    Add_WallBack v1, Vector(v2.x, v2.y, v1.z), Tu, TV, TextureIDBack

    Add_WallFront Vector(v1.x, v1.y, v2.z), Vector(v2.x, v2.y, v2.z), Tu, TV, TextureIDFront

    Add_WallLeft Vector(v2.x, v1.y, v1.z), v2, Tu, TV, TextureIDLeft

    Add_WallRight v1, Vector(v1.x, v2.y, v2.z), Tu, TV, TextureIDRight

    Add_WallFloor Vector(v1.x, v2.y, v1.z), Vector(v2.x, v2.y, v2.z), Tu, TV, TextureIDTop

    Add_WallRoof v1, Vector(v2.x, v1.y, v2.z), Tu, TV, TextureIDBottom

End Sub

Sub Add_Cilynder(Center As D3DVECTOR, ByVal Radius As Long, ByVal Height As Long, Optional ByVal Segment As Long = 10, Optional ByVal TextureIDBody As Long)

  Dim DeltaSegAngle
  Dim SegmentLength, CurrentSegMent
  Dim I, X0, Z0, tu0 As Single, tv0 As Single
  Dim V(2) As D3DVECTOR
  Dim PLANE As D3DPLANE

    DeltaSegAngle = (2# * QUEST3D_PI / Segment)
    SegmentLength = 1# / Segment

    For CurrentSegMent = 0 To Segment

        X0 = Radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = Radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        V(0) = Vector(CSng(X0), Center.y, CSng(Z0))
        V(1) = Vector(CSng(X0), Center.y + Height, CSng(Z0))
        V(2) = Vector(CSng(X0), Center.y, CSng(Z0))

        D3DXPlaneFromPoints PLANE, V(0), V(1), V(2)

        Add_Vertex CSng(X0), Center.y, CSng(Z0), PLANE.A, PLANE.b, PLANE.c, 1# - SegmentLength * (CurrentSegMent + 1), 0, TextureIDBody

        X0 = Radius * Sin(CurrentSegMent * DeltaSegAngle) + Center.x
        Z0 = Radius * Cos(CurrentSegMent * DeltaSegAngle) + Center.z

        Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), PLANE.A, PLANE.b, PLANE.c, 1# - SegmentLength * (CurrentSegMent), 1, TextureIDBody

        Add_Vertex CSng(X0), Center.y, CSng(Z0), PLANE.A, PLANE.b, PLANE.c, 1# - SegmentLength * (CurrentSegMent), 0, TextureIDBody

        X0 = Radius * Sin(CurrentSegMent * DeltaSegAngle) + Center.x
        Z0 = Radius * Cos(CurrentSegMent * DeltaSegAngle) + Center.z

        V(0) = Vector(CSng(X0), Center.y + Height, CSng(Z0))
        V(1) = Vector(CSng(X0), Center.y, CSng(Z0))
        V(2) = Vector(CSng(X0), Center.y + Height, CSng(Z0))

        D3DXPlaneFromPoints PLANE, V(0), V(1), V(2)

        Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), PLANE.A, PLANE.b, PLANE.c, 1# - SegmentLength * (CurrentSegMent), 1, TextureIDBody

        X0 = Radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = Radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        Add_Vertex CSng(X0), Center.y, CSng(Z0), PLANE.A, PLANE.b, PLANE.c, 1# - SegmentLength * (CurrentSegMent + 1), 0, TextureIDBody

        X0 = Radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = Radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), PLANE.A, PLANE.b, PLANE.c, 1# - SegmentLength * (CurrentSegMent + 1), 1, TextureIDBody

    Next CurrentSegMent

    'add top

    For CurrentSegMent = 0 To Segment

        X0 = Radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = Radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        tu0 = (0.5 * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle)) + 0.5
        tv0 = (0.5 * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle)) + 0.5

        Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), 0, 1, 0, tu0, tv0, TextureIDBody

        Add_Vertex Center.x, Center.y + Height, Center.z, 0, 1, tu0, tv0, 0, TextureIDBody

        X0 = Radius * Sin(CurrentSegMent * DeltaSegAngle) + Center.x
        Z0 = Radius * Cos(CurrentSegMent * DeltaSegAngle) + Center.z

        tu0 = (0.5 * Sin(CurrentSegMent * DeltaSegAngle)) + 0.5
        tv0 = (0.5 * Cos(CurrentSegMent * DeltaSegAngle)) + 0.5

        Add_Vertex CSng(X0), Center.y + Height, CSng(Z0), 0, 1, 0, tu0, tv0, TextureIDBody

    Next CurrentSegMent

    'add top bottom

    For CurrentSegMent = 0 To Segment

        X0 = Radius * Sin(CurrentSegMent * DeltaSegAngle) + Center.x
        Z0 = Radius * Cos(CurrentSegMent * DeltaSegAngle) + Center.z

        tu0 = (0.5 * Sin(CurrentSegMent * DeltaSegAngle)) + 0.5
        tv0 = (0.5 * Cos(CurrentSegMent * DeltaSegAngle)) + 0.5

        Add_Vertex CSng(X0), Center.y, CSng(Z0), 0, 1, 0, tu0, tv0, TextureIDBody

        Add_Vertex Center.x, Center.y, Center.z, 0, 1, tu0, tv0, 0, TextureIDBody

        X0 = Radius * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.x
        Z0 = Radius * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle) + Center.z

        tu0 = (0.5 * Sin(CurrentSegMent * DeltaSegAngle + DeltaSegAngle)) + 0.5
        tv0 = (0.5 * Cos(CurrentSegMent * DeltaSegAngle + DeltaSegAngle)) + 0.5

        Add_Vertex CSng(X0), Center.y, CSng(Z0), 0, 1, 0, tu0, tv0, TextureIDBody

    Next CurrentSegMent

End Sub

Sub Add_WallBack(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional ByVal Tu As Single = 1, Optional ByVal TV As Single = 1, Optional ByVal TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

  Dim V(3) As D3DVECTOR

    V(0) = Vector(v1.x, v1.y, v1.z)
    V(1) = Vector(v1.x, v2.y, v1.z)
    V(2) = Vector(v2.x, v1.y, v2.z)

  Dim P As D3DPLANE

    D3DXPlaneFromPoints P, V(0), V(1), V(2)

    Call Add_Vertex(v1.x, v1.y, v1.z, P.A, P.b, P.c, 0, TV, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, P.A, P.b, P.c, 0, 0, TextureID)
    Call Add_Vertex(v2.x, v1.y, v2.z, P.A, P.b, P.c, Tu, TV, TextureID)

    V(0) = Vector(v1.x, v2.y, v1.z)
    V(1) = Vector(v2.x, v2.y, v2.z)
    V(2) = Vector(v2.x, v1.y, v1.z)

    D3DXPlaneFromPoints P, V(0), V(1), V(2)

    Call Add_Vertex(v1.x, v2.y, v1.z, P.A, P.b, P.c, 0, 0, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, P.A, P.b, P.c, Tu, 0, TextureID)
    Call Add_Vertex(v2.x, v1.y, v1.z, P.A, P.b, P.c, Tu, TV, TextureID)

End Sub

'----------------------------------------
'Name: Add_WallFloor
'Object: Add
'Event: WallFloor
'----------------------------------------
Sub Add_WallFloor(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional ByVal Tu As Single = 1, Optional ByVal TV As Single = 1, Optional ByVal TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    Call Add_Vertex(v1.x, v2.y, v2.z, 0, 1, 0, 0, 0, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, 1, 0, Tu, 0, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, 1, 0, 0, TV, TextureID)

    Call Add_Vertex(v2.x, v2.y, v1.z, 0, 1, 0, Tu, TV, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, 1, 0, 0, TV, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, 1, 0, Tu, 0, TextureID)

End Sub

'----------------------------------------
'Name: Add_WallFront
'Object: Add
'Event: WallFront
'----------------------------------------
Sub Add_WallFront(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional ByVal Tu As Single = 1, Optional ByVal TV As Single = 1, Optional ByVal TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

  Dim V(3) As D3DVECTOR

    V(0) = Vector(v2.x, v1.y, v2.z)
    V(1) = Vector(v1.x, v2.y, v1.z)
    V(2) = Vector(v1.x, v1.y, v1.z)

  Dim P As D3DPLANE

    D3DXPlaneFromPoints P, V(0), V(1), V(2)

    Call Add_Vertex(v2.x, v1.y, v2.z, P.A, P.b, P.c, Tu, TV, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, P.A, P.b, P.c, 0, 0, TextureID)
    Call Add_Vertex(v1.x, v1.y, v1.z, P.A, P.b, P.c, 0, TV, TextureID)

    V(0) = Vector(v2.x, v1.y, v1.z)
    V(1) = Vector(v2.x, v2.y, v2.z)
    V(2) = Vector(v1.x, v2.y, v1.z)

    D3DXPlaneFromPoints P, V(0), V(1), V(2)

    '_________
    Call Add_Vertex(v2.x, v1.y, v1.z, P.A, P.b, P.c, Tu, TV, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, P.A, P.b, P.c, Tu, 0, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, P.A, P.b, P.c, 0, 0, TextureID)

End Sub

'----------------------------------------
'Name: Add_WallLeft
'Object: Add
'Event: WallLeft
'----------------------------------------
Sub Add_WallLeft(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional ByVal Tu As Single = 1, Optional ByVal TV As Single = 1, Optional ByVal TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight
  Dim V(3) As D3DVECTOR

    V(0) = Vector(v1.x, v1.y, v1.z)
    V(1) = Vector(v1.x, v2.y, v1.z)
    V(2) = Vector(v1.x, v1.y, v2.z)
    V(3) = Vector(v1.x, v2.y, v2.z)

    'V(0).Tu = 0: V(0).TV = 1
    'V(1).Tu = 0: V(1).TV = 0
    'V(2).Tu = 1: V(2).TV = 1
    'V(3).Tu = 1: V(3).TV = 0
    'V(0).Tu = 0: V(0).TV = 1
    'V(1).Tu = 0: V(1).TV = 0
    'V(2).Tu = 1: V(2).TV = 1
    'V(3).Tu = 1: V(3).TV = 0

  Dim P As D3DPLANE

    D3DXPlaneFromPoints P, V(0), V(1), V(2)

    Add_VertexEx V(0), Vector(P.A, P.b, P.c), Vector2D(0, TV), TextureID
    Add_VertexEx V(1), Vector(P.A, P.b, P.c), Vector2D(0, 0), TextureID
    Add_VertexEx V(2), Vector(P.A, P.b, P.c), Vector2D(Tu, TV), TextureID

    D3DXPlaneFromPoints P, V(1), V(3), V(2)

    Add_VertexEx V(1), Vector(P.A, P.b, P.c), Vector2D(0, 0), TextureID
    Add_VertexEx V(3), Vector(P.A, P.b, P.c), Vector2D(Tu, 0), TextureID
    Add_VertexEx V(2), Vector(P.A, P.b, P.c), Vector2D(Tu, TV), TextureID

End Sub

'----------------------------------------
'Name: Add_WallRight
'Object: Add
'Event: WallRight
'----------------------------------------
Sub Add_WallRight(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional ByVal Tu As Single = 1, Optional ByVal TV As Single = 1, Optional ByVal TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight
  Dim V(3) As D3DVECTOR

    V(0) = Vector(v1.x, v1.y, v2.z)
    V(1) = Vector(v1.x, v2.y, v2.z)
    V(2) = Vector(v1.x, v1.y, v1.z)
    V(3) = Vector(v1.x, v2.y, v1.z)
    'V(0).Tu = 0: V(0).TV = 1
    'V(1).Tu = 0: V(1).TV = 0
    'V(2).Tu = 1: V(2).TV = 1
    'V(3).Tu = 1: V(3).TV = 0

  Dim P As D3DPLANE

    D3DXPlaneFromPoints P, V(0), V(1), V(2)

    Add_VertexEx V(0), Vector(P.A, P.b, P.c), Vector2D(0, TV), TextureID
    Add_VertexEx V(1), Vector(P.A, P.b, P.c), Vector2D(0, 0), TextureID
    Add_VertexEx V(2), Vector(P.A, P.b, P.c), Vector2D(Tu, TV), TextureID

    D3DXPlaneFromPoints P, V(1), V(3), V(2)

    Add_VertexEx V(1), Vector(P.A, P.b, P.c), Vector2D(0, 0), TextureID
    Add_VertexEx V(3), Vector(P.A, P.b, P.c), Vector2D(Tu, 0), TextureID
    Add_VertexEx V(2), Vector(P.A, P.b, P.c), Vector2D(Tu, TV), TextureID

End Sub

'----------------------------------------
'Name: Add_WallRoof
'Object: Add
'Event: WallRoof
'----------------------------------------
Sub Add_WallRoof(VBotomLeft As D3DVECTOR, VTopRight As D3DVECTOR, Optional ByVal Tu As Single = 1, Optional ByVal TV As Single = 1, Optional ByVal TextureID As Long = 0)

  Dim v1 As D3DVECTOR
  Dim v2 As D3DVECTOR

    v1 = VBotomLeft
    v2 = VTopRight

    Call Add_Vertex(v2.x, v2.y, v2.z, 0, -1, 0, Tu, 0, TextureID)
    Call Add_Vertex(v1.x, v2.y, v1.z, 0, -1, 0, 0, TV, TextureID)
    Call Add_Vertex(v2.x, v2.y, v1.z, 0, -1, 0, Tu, TV, TextureID)

    Call Add_Vertex(v1.x, v2.y, v1.z, 0, -1, 0, 0, TV, TextureID)
    Call Add_Vertex(v2.x, v2.y, v2.z, 0, -1, 0, Tu, 0, TextureID)
    Call Add_Vertex(v1.x, v2.y, v2.z, 0, -1, 0, 0, 0, TextureID)

End Sub

Sub BuildMesh()

  Dim I, J, z
  Dim VertexSizeInBytes As Long
  Dim VV As QUEST3D_VERTEX2

    VertexSizeInBytes = Len(VV)

    With mMESH

  Dim TT As Long

        RaiseEvent Progress("Optimize Polygons", 0)

        Erase mMESH.Vertex

  Dim t() As QUEST3D_VERTEX2
  Dim N As Long
  Dim V As Long

        '  If .Header.NumText = 1 Then
        '    If .POLY_sorted(0).NumVERT <= 6 Then
        '
        '      Dim R As Single
        '      R = MAX_3(.SUB_BOX(0).Vmax.x - .SUB_BOX(0).Vmin.x, _
            '      .SUB_BOX(0).Vmax.y - .SUB_BOX(0).Vmin.y, .SUB_BOX(0).Vmax.Z - .SUB_BOX(0).Vmin.Z)
        '
        '      .SUB_BOX(0).Vmin.x = .SUB_BOX(0).Vmin.x - R / 3
        '      .SUB_BOX(0).Vmin.y = .SUB_BOX(0).Vmin.y - R / 3
        '      .SUB_BOX(0).Vmin.Z = .SUB_BOX(0).Vmin.Z - R / 3
        '
        '
        '     .SUB_BOX(0).Vmax.x = .SUB_BOX(0).Vmax.x + R / 3
        '      .SUB_BOX(0).Vmax.y = .SUB_BOX(0).Vmax.y + R / 3
        '      .SUB_BOX(0).Vmax.Z = .SUB_BOX(0).Vmax.Z + R / 3
        '
        '    End If
        '  End If

        For I = 0 To .Header.NumText - 1
            On Error Resume Next

                RaiseEvent Progress("Computes Normals and Optimize meshes", Int((I / (.Header.NumText - 1)) * 100))
                Dim K(10) As Boolean

                Fill_Mesh .POLY_sorted(I).Vert, .POLY_sorted(I).DXMESH, .POLY_sorted(I).DXMESH2, .POLY_sorted(I).DXMESH3, .POLY_sorted(I).DXMESH4, .POLY_sorted(I).DXMESH5, _
                .POLY_sorted(I).DXMESH6, .POLY_sorted(I).DXMESH7, .POLY_sorted(I).DXMESH8, .POLY_sorted(I).DXMESH9, .POLY_sorted(I).DXMESH10, K

                If K(1) = True Then .POLY_sorted(I).DXMESH_FAIL = True
                If K(2) = True Then .POLY_sorted(I).DXMESH_FAIL2 = True
                If K(3) = True Then .POLY_sorted(I).DXMESH_FAIL3 = True
                If K(4) = True Then .POLY_sorted(I).DXMESH_FAIL4 = True
                If K(5) = True Then .POLY_sorted(I).DXMESH_FAIL5 = True
                If K(6) = True Then .POLY_sorted(I).DXMESH_FAIL6 = True
                If K(7) = True Then .POLY_sorted(I).DXMESH_FAIL7 = True
                If K(8) = True Then .POLY_sorted(I).DXMESH_FAIL8 = True
                If K(9) = True Then .POLY_sorted(I).DXMESH_FAIL9 = True
                
                

                '                Init.DXmeshFromVERT2 .POLY_sorted(I).DXMESH, .POLY_sorted(I).VERT
                '
                '                .POLY_sorted(I).DXMESH_FAIL = (.POLY_sorted(I).DXMESH Is Nothing)
                '
                '                If .POLY_sorted(I).DXMESH_FAIL Then
                '
                '                    n = (.POLY_sorted(I).TotalVERT / 3) / 2
                '
                '                    If n / 2 Mod 2 = 0 Then
                '                        n = n
                '                      Else
                '                        n = n - 1
                '                    End If
                '                    n = n * 3
                '
                '                    ReDim t(n - 1)
                '                    CopyMemory t(0), .POLY_sorted(I).VERT(0), 40 * n
                '                    Init.DXmeshFromVERT2 .POLY_sorted(I).DXMESH, t
                '
                '                    V = n
                '
                '                    n = .POLY_sorted(I).TotalVERT - n
                '                    ReDim t(n - 1)
                '                    CopyMemory t(0), .POLY_sorted(I).VERT(V), 40 * n
                '                    Init.DXmeshFromVERT2 .POLY_sorted(I).DXMESH2, t
                '
                '                End If
                '
                '                .POLY_sorted(I).DXMESH_FAIL2 = (.POLY_sorted(I).DXMESH Is Nothing) And (.POLY_sorted(I).DXMESH2 Is Nothing)
                '
                '                If .POLY_sorted(I).DXMESH_FAIL2 Then
                '
                '                    n = (.POLY_sorted(I).TotalVERT / 3) / 3
                '
                '                    If n / 2 Mod 2 = 0 Then
                '                        n = n
                '                      Else
                '                        n = n - 1
                '                    End If
                '                    n = n * 3
                '
                '                    ReDim t(n - 1)
                '                    CopyMemory t(0), .POLY_sorted(I).VERT(0), 40 * n
                '                    Init.DXmeshFromVERT2 .POLY_sorted(I).DXMESH, t
                '
                '                    V = n
                '
                '                    CopyMemory t(0), .POLY_sorted(I).VERT(0), 40 * n
                '                    Init.DXmeshFromVERT2 .POLY_sorted(I).DXMESH2, t
                '
                '                    n = .POLY_sorted(I).TotalVERT - n
                '                    ReDim t(n - 1)
                '                    CopyMemory t(0), .POLY_sorted(I).VERT(V), 40 * n
                '                    Init.DXmeshFromVERT2 .POLY_sorted(I).DXMESH3, t
                '
                '                End If

                '                obj_d3dx.ComputeNormals .POLY_sorted(I).DXMESH
                '
                '                If .POLY_sorted(I).DXMESH_FAIL Then
                '                    obj_d3dx.ComputeNormals .POLY_sorted(I).DXMESH2
                '                End If
                '
                '                If .POLY_sorted(I).DXMESH_FAIL2 Then
                '                    obj_d3dx.ComputeNormals .POLY_sorted(I).DXMESH3
                '                End If
                '
                '                If .POLY_sorted(I).DXMESH_FAIL3 Then
                '                    obj_d3dx.ComputeNormals .POLY_sorted(I).DXMESH4
                '                End If
                '
                '                If .POLY_sorted(I).DXMESH_FAIL4 Then
                '                    obj_d3dx.ComputeNormals .POLY_sorted(I).DXMESH5
                '                End If

                Set .POLY_sorted(I).VBUF = .POLY_sorted(I).DXMESH.GetVertexBuffer
                Set .POLY_sorted(I).iBuf = .POLY_sorted(I).DXMESH.GetIndexBuffer

                '            Device.SetStreamSource 0, vBuf, obj_d3dx.GetFVFVertexSize(D3DFVF_VERTEX)
                '            Device.SetIndices iBuf, 0
                '
                .POLY_sorted(I).nIndici = .POLY_sorted(I).DXMESH.GetNumVertices
                .POLY_sorted(I).nFacce = .POLY_sorted(I).DXMESH.GetNumFaces

                .POLY_sorted(I).SizeofVert = obj_D3DX.GetFVFVertexSize(mMESH.FVF)

                'End If
                .NumP = .NumP + .POLY_sorted(I).TotalVERT / 3
            Next I

        End With

        Me.Set_Position 0, 0, 0
        Me.Set_Rotation 0, 0, 0
        Me.Set_Scale 1, 1, 1

End Sub

Function Add_Texture(ByVal TextureFile As String) As Integer

  'If FileiS_valid(TextureFile) Then

    mMESH.Header.NumText = mMESH.Header.NumText + 1
    ReDim Preserve mMESH.TexturesID(mMESH.Header.NumText - 1)

    ReDim Preserve mMESH.TexturesType(mMESH.Header.NumText - 1)

    ReDim Preserve mMESH.TextureInfo(mMESH.Header.NumText - 1)

    ReDim Preserve mMESH.POLY_sorted(mMESH.Header.NumText - 1)

    mMESH.TexturesType(mMESH.Header.NumText - 1) = OPAQUE

  Dim tempStr As String
    tempStr = Get_fileNameEX(TextureFile)
    TexturePATH = Get_pathName(TextureFile)

  
    mMESH.TextureInfo(mMESH.Header.NumText - 1).Filename = tempStr

   
    mMESH.TexturesID(mMESH.Header.NumText - 1) = Add_TextureToPool(TextureFile, , , 0)

    Add_Texture = -1
    If mMESH.TexturesID(mMESH.Header.NumText - 1) > -1 Then
        Add_Texture = mMESH.Header.NumText - 1
    End If

End Function

Function Add_Texture2(ByVal TextureFile As String, Optional ByVal ColorKeyARGB As Long = -1) As Long

  Dim Texture As Direct3DBaseTexture8

    Set Texture = CreateTextureColorKEY(TextureFile, , , ColorKeyARGB)

    mMESH.Header.NumText = mMESH.Header.NumText + 1
    ReDim Preserve mMESH.TexturesID(mMESH.Header.NumText - 1)
    ReDim Preserve mMESH.TexturesType(mMESH.Header.NumText - 1)

  
    ReDim Preserve mMESH.TextureInfo(mMESH.Header.NumText - 1)
   

    ReDim Preserve mMESH.POLY_sorted(mMESH.Header.NumText - 1)

  
    mMESH.TexturesType(mMESH.Header.NumText - 1) = TRANSPARENT

   
   
    mMESH.TexturesID(mMESH.Header.NumText - 1) = Add_TextureToPoolEX(Texture)

   
    If mMESH.TexturesID(mMESH.Header.NumText - 1) > -1 Then
        Add_Texture2 = mMESH.Header.NumText - 1
    End If

End Function

Sub Set_Position(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    mMESH.Xdata.vPosition.x = x
    mMESH.Xdata.vPosition.y = y
    mMESH.Xdata.vPosition.z = z
    mMESH.Xdata.HasChanged = 1
    Compute

End Sub

Sub Set_PositionEx(vPos As D3DVECTOR)

    Set_Position vPos.x, vPos.y, vPos.z

End Sub

Function Get_Position() As D3DVECTOR

    Get_Position = mMESH.Xdata.vPosition

End Function

Function Get_Rotation() As D3DVECTOR

    Get_Rotation = mMESH.Xdata.Vrotate

End Function

Function Get_Scale() As D3DVECTOR

    Get_Scale = mMESH.Xdata.Vscal

End Function

'----------------------------------------
'Name: Set_Rotation
'Object: Set
'Event: Rotation
'Description:
'----------------------------------------
Sub Set_Rotation(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    mMESH.Xdata.Vrotate.x = x
    mMESH.Xdata.Vrotate.y = y
    mMESH.Xdata.Vrotate.z = z
    mMESH.Xdata.HasChanged = 1
    Compute

End Sub

'----------------------------------------
'Name: Set_Scale
'Object: Set
'Event: Scale
'Description:
'----------------------------------------
Sub Set_Scale(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    mMESH.Xdata.Vscal.x = x
    mMESH.Xdata.Vscal.y = y
    mMESH.Xdata.Vscal.z = z
    mMESH.Xdata.HasChanged = 1
    Compute

End Sub

Private Sub Compute()

    If mMESH.Xdata.HasChanged Then

        mMESH.Xdata.HasChanged = 0

        mMESH.Xdata.WorldMatrix = Matrix_GetEX(mMESH.Xdata.Vscal, mMESH.Xdata.Vrotate, mMESH.Xdata.vPosition)

        D3DXMatrixInverse mMESH.Xdata.WorldInvMatrix, 0, mMESH.Xdata.WorldMatrix

    End If

End Sub

Sub Render()

  Dim I As Long
  Dim KK As Long 'for boxes collided list

    'Dim BB() As Long
  Dim J As Long
  Dim lpNod As Long

  Dim TempVERTS As Long

    'we save renderstates

    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_LIGHTING

    LpGLOBAL_QUEST3D.Set_EngineLight True

    obj_Device.SetTransform D3DTS_WORLD, mMESH.Xdata.WorldMatrix

    obj_Device.SetVertexShader mMESH.FVF

    For I = 0 To mMESH.Header.NumText - 1
        With mMESH

            TempVERTS = .POLY_sorted(I).TotalVERT

            If TempVERTS = 0 Then GoTo NEXT_SUBMESH

            If .TexturesID(I) > -1 Then _
               obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(.TexturesID(I))

            If .TexturesType(I) = TRANSPARENT Then

                LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ALPHABLENDENABLE
                LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ALPHATESTENABLE
                LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ALPHAREF
                LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ALPHAFUNC
                LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_DESTBLEND
                LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_DESTBLEND

                obj_Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1  'TRUE
                obj_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                obj_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

                ' Enable alpha testing (skips pixels with less than a certain alpha.)
                'If ((d3dCaps.AlphaCmpCaps And D3DPCMPCAPS_GREATEREQUAL) = D3DPCMPCAPS_GREATEREQUAL) Then
                obj_Device.SetRenderState D3DRS_ALPHAREF, &HF0&

                obj_Device.SetRenderState D3DRS_ALPHATESTENABLE, 1 'TRUE
                obj_Device.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL

                '           If .TexturesID(i) > -1 Then _
                    '               obj_device.SetTexture 0, myTEXPOOL.POOL_texture(.TexturesID(i)) Else obj_device.SetTexture 0, Nothing
                '
                obj_Device.DrawPrimitiveUP D3DPT_TRIANGLELIST, Int(.POLY_sorted(I).NumVert / 3), .POLY_sorted(I).Vert(0), 40

                LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ALPHABLENDENABLE
                LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ALPHATESTENABLE
                LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ALPHAREF
                LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ALPHAFUNC
                LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_DESTBLEND
                LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_DESTBLEND

              Else
              '=========================================================================
              ' Here we check if this sub mesh has other chunk than the primary
              '
              '
              '========================================================================

                .POLY_sorted(I).DXMESH.DrawSubset 0

                'chunk N°2 has been found then we draw it
                'via DxMesh
                If .POLY_sorted(I).DXMESH_FAIL Then _
                   .POLY_sorted(I).DXMESH2.DrawSubset 0

                'chunk N°3 has been found then we draw it
                'via DxMesh
                If .POLY_sorted(I).DXMESH_FAIL2 Then _
                   .POLY_sorted(I).DXMESH3.DrawSubset 0
                
                'chunk N°4 has been found then we draw it
                'via DxMesh
                If .POLY_sorted(I).DXMESH_FAIL3 Then _
                   .POLY_sorted(I).DXMESH4.DrawSubset 0

                '..................IDEM..................
                If .POLY_sorted(I).DXMESH_FAIL4 Then _
                   .POLY_sorted(I).DXMESH5.DrawSubset 0
                   
                '..................IDEM..................
                 If .POLY_sorted(I).DXMESH_FAIL5 Then _
                   .POLY_sorted(I).DXMESH6.DrawSubset 0
                   
                '..................IDEM..................
                 If .POLY_sorted(I).DXMESH_FAIL6 Then _
                   .POLY_sorted(I).DXMESH7.DrawSubset 0


                '..................IDEM..................
                If .POLY_sorted(I).DXMESH_FAIL7 Then _
                   .POLY_sorted(I).DXMESH8.DrawSubset 0

                '..................IDEM..................
                If .POLY_sorted(I).DXMESH_FAIL8 Then _
                   .POLY_sorted(I).DXMESH9.DrawSubset 0
                   
                '..................IDEM..................
                 If .POLY_sorted(I).DXMESH_FAIL9 Then _
                   .POLY_sorted(I).DXMESH10.DrawSubset 0
                   
            End If

            'we tell the engine that verices
            'and polygons have been pushed
            Add_Tri TempVERTS / 3
            Add_Verti TempVERTS

        End With

        obj_Device.SetTexture 0, Nothing

NEXT_SUBMESH:
    Next I

    'recall previous renderstates

    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_LIGHTING

End Sub

Private Sub Class_Initialize()

    mMESH.FVF = QUEST3D_FVFVERTEX2

End Sub



Sub Free()

  Dim I As Long

    If mMESH.Header.IS_LOADED And mMESH.Header.IS_Builded Then
        mMESH.Header.IS_LOADED = 0
        mMESH.Header.IS_Builded = 0
       
        
        Erase mMESH.TextureInfo
        Erase mMESH.TextureInfo
        Erase mMESH.POLY_sorted
        mMESH.Header.IS_LOADED = 0
        mMESH.Header.IS_Builded = 0
        ReDim mMESH.POLY_sorted(0)

        For I = 0 To UBound(mMESH.POLY_sorted)

            Set mMESH.POLY_sorted(I).DXMESH = Nothing

            If mMESH.POLY_sorted(I).DXMESH_FAIL Then _
               Set mMESH.POLY_sorted(I).DXMESH2 = Nothing

            If mMESH.POLY_sorted(I).DXMESH_FAIL2 Then _
               Set mMESH.POLY_sorted(I).DXMESH3 = Nothing

            If mMESH.POLY_sorted(I).DXMESH_FAIL3 Then _
               Set mMESH.POLY_sorted(I).DXMESH4 = Nothing

            If mMESH.POLY_sorted(I).DXMESH_FAIL4 Then _
               Set mMESH.POLY_sorted(I).DXMESH5 = Nothing
               
            
            
             If mMESH.POLY_sorted(I).DXMESH_FAIL5 Then _
               Set mMESH.POLY_sorted(I).DXMESH6 = Nothing

            If mMESH.POLY_sorted(I).DXMESH_FAIL6 Then _
               Set mMESH.POLY_sorted(I).DXMESH7 = Nothing

            If mMESH.POLY_sorted(I).DXMESH_FAIL8 Then _
               Set mMESH.POLY_sorted(I).DXMESH9 = Nothing

            If mMESH.POLY_sorted(I).DXMESH_FAIL9 Then _
               Set mMESH.POLY_sorted(I).DXMESH10 = Nothing
               
               
               
               

            Set mMESH.POLY_sorted(I).VBUF = Nothing
            Set mMESH.POLY_sorted(I).iBuf = Nothing

        Next I
        Erase mMESH.POLY_sorted
        ReDim mMESH.POLY_sorted(0)
        mMESH.Header.IS_LOADED = 0

    End If

End Sub

