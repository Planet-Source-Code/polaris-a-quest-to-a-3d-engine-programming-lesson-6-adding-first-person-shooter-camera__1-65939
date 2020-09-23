VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Engine_Lesson 6 Adding First Person Shooter Camera"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================
'WELCOME Engine_Lesson 6 Adding First Person Shooter Camera
'_____________________________________________________________________________________
'-------------------------------------------------------------------------------------
'
'===================================================================================
'Welcome to this Step by step Quest to a 3D Engine programming
'this tutorial will show you how to design a simple 3D
'engine, Next tutorials will show how to add other engine objet
'like Camera,Mesh and Object Polygon
'
'This tutorial 6: Adding First Person Shooter Camera
'
'It shows you how to
'  - Add an advanced camera system go at cQuest_Camera class for code
'  - use direct Input See Module_Input
'  - We use advanced settings to improve rendering speed and quality
'
'
'

'How to read this code
'   - Form1: is the engine code in action
'   - frmEnum have code for Device anumeration
'   - Module_definitions will hold all engine objets definitions and types
'   - Module_Util will hold all Vector,matrix,Color math stuff
'   - cQuest3D_Core is our first object, it defines Main entry of the engine
'   - cQuest3D_Mesh 3D Mesh Class, to generate procedurally 3D Scene
'   - cQuest3D_Input Define Input class (Keyboard,Mouse,Joystick)
'   - cQuest_Camera handle Camera

'
'Good coding
'
'Vote if you want the sequel!!
'
'==================================================================================

Option Explicit

'we use the engine here
'we declare an objet
Dim QUEST As cQuest3D_Core

'mesh
Dim MyMESH1 As cQuest3D_Mesh
Dim MyMESH2 As cQuest3D_Mesh

'camera
Dim KAMERA As cQuest3D_Camera
'we use these vars to do
'frame based camera animation
Const ROTATION_SPEED As Single = 1
Const PLAYER_SPEED As Single = 500
'see GetInput()

'Input
Dim KEY As cQuest3D_Input
Dim ShowInfo As Boolean

Private Sub Form_Load()

    InitEngine

End Sub

Sub InitEngine()

  'we allocate memory here

    Set QUEST = New cQuest3D_Core

    '        'we initialize the engine
    '        If QUEST.Init_Dialogue(Me.hwnd) = False Then
    '         MsgBox "Sorry there was an error"
    '         End
    '        End If

    If QUEST.Init(Me.hwnd) = False Then
        MsgBox "Sorry there was an error"
        End
    End If

    Me.Refresh
    Me.Show

    'init default state for the engine
    QUEST.Set_EngineAmbientColor Make_ColorRGB(255, 255, 255) 'white

    QUEST.Set_BackbufferClearColor Make_ColorRGB(100, 100, 100)

    'we want to clear zbuffer and backbuffer
    QUEST.Set_EngineClearRenderTarget True
    'fill mode
    QUEST.Set_EngineFillMode QUEST3D_FILL_SOLID

    'shade mode
    QUEST.Set_EngineShadeMode QUEST3D_SHADE_GOURAUD
    'choose from:
    '    QUEST3D_FILTER_POINT = 0
    '    QUEST3D_FILTER_BILINEAR = 1
    '    QUEST3D_FILTER_TRILINEAR = 2
    '    QUEST3D_FILTER_ANISOTROPIC = 3
    ' Best is QUEST3D_FILTER_ANISOTROPIC
    QUEST.Set_EngineTextureFilter QUEST3D_FILTER_TRILINEAR

    'here we define mimap filter
    'what is a mipmap?
    'A mipmap is reduce sized textures that allow
    'texture mapping on small polygon size,it allow to preserve drawing details
    'we choos the best filter to apply to these mipmaps
    'because the engine we are building tells to Direct3D to generate mipmaps
    'for each texture we use DEFAULT filtering method
    QUEST.Set_EngineTextureMipMapFilter QUEST3D_TEXTURE_DEFAULT

    'prepare input

    Set KEY = New cQuest3D_Input
    'we force input devices creation
    KEY.ReCreateInputDevices
    'we force device polling
    KEY.ReCreateInputDevices

    ShowInfo = True

    'init Camera
    Set KAMERA = New cQuest3D_Camera
    'we set the first person shooter mode
    KAMERA.Set_CameraStyle FPS_STYLE
    'we use Left Hand perpective projection
    KAMERA.Set_CameraProjectionType PT_PERSPECTIVE_LH
    'we use A field of view where near=10 and far=10000, Angle=45 degree
    KAMERA.Set_ViewFrustum 10, 10000, 45 * QUEST3D_RAD
    'initial camera position=0,100,0 looking at 0,100,100
    KAMERA.Set_camera Vector(0, 100, 0), Vector(0, 100, 100)
    'we update the camera
    KAMERA.Update

    'we prapare geometry
    PrepareGeometry
    'we call game loop
    GameLoop

End Sub

'==================================================================================
'
'In this sub we procedurally generate
'geometry here, the floor and cylinders
'
'==================================================================================

Sub PrepareGeometry()

  Dim Texture_ID As Long
  Dim I As Long

    'here we allocate memory for mesh
    Set MyMESH1 = New cQuest3D_Mesh

    '1st we add textures
    Texture_ID = MyMESH1.Add_Texture(App.Path + "\Data\castle_m04.jpg") '(ID 1)

    '2nd we ad vertices,polygons
    'MyMESH1.Add_WallFloor Vector(-9000, -1, -9000), Vector(9000, -1, 9000), 10, 10, 0 '0 means we used fisrt textures added

    'we add randomly center cylinders
    For I = 1 To 10 '80 can be changed to 850 max ,This Engine can handle 120 000 Polygons Max per Sub mesh

        MyMESH1.Add_Cilynder Vector((Rnd - Rnd) * 10000, -1, (Rnd - Rnd) * 10000), 100 + (Rnd - Rnd) * 50 + 50, 500 + (Rnd - Rnd) * 200 + 100, 10 + (Rnd) * 40, Texture_ID '0 means 1st texture added

    Next I

    '3rd we Build the mesh
    'all information for fast rendering will be
    'computed
    MyMESH1.BuildMesh

    'now we define a more complex scene

    'here we allocate memory for mesh
    Set MyMESH2 = New cQuest3D_Mesh

    MyMESH2.Add_Texture (App.Path + "\Data\Relief_8.jpg")          '0
    MyMESH2.Add_Texture (App.Path + "\Data\facade2.jpg")  '1

    MyMESH2.Add_Texture App.Path + "\Data\street.JPG"     '2
    MyMESH2.Add_Texture App.Path + "\Data\street2.JPG"    '3
    MyMESH2.Add_Texture (App.Path + "\Data\kerb2.jpg")    '4
    MyMESH2.Add_Texture (App.Path + "\Data\kerb.jpg")     '5

    MyMESH2.Add_Texture (App.Path + "\Data\square.JPG")   '6

    MyMESH2.Add_Texture App.Path + "\Data\cement.JPG"     '7

    MyMESH2.Add_Texture App.Path + "\Data\road_t03.jpg"   '8

    MyMESH2.Add_Texture App.Path + "\Data\Asfalto1.bmp"   '9
    MyMESH2.Add_Texture App.Path + "\Data\pierres.JPG"    '10

    MyMESH2.Add_Texture App.Path + "\Data\win2.JPG"       '11
    MyMESH2.Add_Texture App.Path + "\Data\windows.JPG"       '12

    MyMESH2.Add_Texture App.Path + "\Data\StoreSd.BMP"     '13

    'city 1 floor
    MyMESH2.Add_WallFloor Vector(-5000, -1, -5000), Vector(5000, -1, 5000), 10, 10, 0

    'city 2 floor
    MyMESH2.Add_WallFloor Vector(-5000, -1, 5000), Vector(5000, -1, 10000), 10, 10, 7

    '=======NEW CODE======='
    'add building 1
    MyMESH2.Add_Box Vector(-500, 0, 0), Vector(0, 500, 500), 1, 1, 1, 1, 1, 1

    'add building 2
    MyMESH2.Add_Box Vector(-500, 0, -2000), Vector(0, 500, -1500), 1, 1, 1, 1, 1, 1

    'add building 3
    MyMESH2.Add_Box Vector(601, 0, 0), Vector(1051, 800, 500), 11, 11, 11, 11, 11, 11

    'add building 4
    MyMESH2.Add_Box Vector(601, 0, -2000), Vector(1051, 900, -1500), 12, 12, 12, 12, 12, 12

    'add building 4
    MyMESH2.Add_Box Vector(601, 0, 2000), Vector(1051, 900, 2500), 13, 13, 13, 13, 13, 13

    'draw the road  segment of 1000 from -5000 to 5000
    MyMESH2.Add_WallFloor Vector(50, 0, -5000), Vector(550, 0, 5000), 1, 2, 2 'the 3rd texture passed to the mesh class

    'draw the north and south roads
    'north part
    MyMESH2.Add_WallFloor Vector(50, 0, 5000), Vector(550, 0, 5500), 1, 1, 6
    MyMESH2.Add_WallFloor Vector(550, 0, 5000), Vector(5000, 0, 5500), 5, 1, 8
    MyMESH2.Add_WallFloor Vector(-5000, 0, 5000), Vector(50, 0, 5500), 5, 1, 8
    'south part
    MyMESH2.Add_WallFloor Vector(550, 0, -5500), Vector(5000, 0, -5000), 5, 1, 8
    MyMESH2.Add_WallFloor Vector(-5000, 0, -5500), Vector(50, 0, -5000), 5, 1, 8
    MyMESH2.Add_WallFloor Vector(50, 0, -5500), Vector(550, 0, -5000), 1, 1, 6

    'make the pavements
    MyMESH2.Add_WallFloor Vector(0, 10, -5000), Vector(50, 10, 5000), 1, 10, 5
    MyMESH2.Add_WallFloor Vector(550, 10, -5000), Vector(600, 10, 5000), 1, 10, 4
    MyMESH2.Add_WallLeft Vector(50, 0, -5000), Vector(50, 10, 5000), 10, 0.25, 9
    MyMESH2.Add_WallRight Vector(550, 0, -5000), Vector(550, 10, 5000), 18, 0.25, 9

    'the south west pavements
    MyMESH2.Add_WallBack Vector(-5000, -1, -5000), Vector(50, 10, -5000), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(-5000, -1, -4950), Vector(0, 10, -4950), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(-5000, 10, -5000), Vector(0, 10, -4950), 18, 1, 9

    MyMESH2.Add_WallBack Vector(-5000, -1, -5500), Vector(50, 10, -5500), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(-5000, -1, -5450), Vector(0, 10, -5450), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(-5000, 10, -5500), Vector(0, 10, -5450), 18, 1, 9

    'the south east pavements
    MyMESH2.Add_WallBack Vector(550, -1, -5000), Vector(5000, 10, -5000), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(550, -1, -4950), Vector(5000, 10, -4950), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(600, 10, -5000), Vector(5000, 10, -4950), 18, 1, 9

    MyMESH2.Add_WallBack Vector(550, -1, -5500), Vector(5000, 10, -5500), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(550, -1, -5450), Vector(5000, 10, -5450), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(500, 10, -5500), Vector(5000, 10, -5450), 18, 1, 9
    'big wall
    MyMESH2.Add_WallFront Vector(-5000, -1, -5501), Vector(5000, 500, -5501), 18, 1, 10

    'the north west pavements
    MyMESH2.Add_WallBack Vector(-5000, -1, 4950), Vector(0, 10, 4950), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(-5000, -1, 5000), Vector(50, 10, 5000), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(-5000, 10, 4950), Vector(0, 10, 5000), 18, 1, 9

    MyMESH2.Add_WallFront Vector(-5000, -1, 5550), Vector(50, 10, 5550), 18, 0.5, 9
    MyMESH2.Add_WallBack Vector(-5000, -1, 5500), Vector(0, 10, 5500), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(-5000, 10, 5500), Vector(0, 10, 5550), 18, 1, 9

    'the north east pavements
    MyMESH2.Add_WallBack Vector(550, -1, 4950), Vector(5000, 10, 4950), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(550, -1, 5000), Vector(5000, 10, 5000), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(600, 10, 4950), Vector(5000, 10, 5000), 18, 1, 9

    MyMESH2.Add_WallFront Vector(550, -1, 5550), Vector(5000, 10, 5550), 18, 0.5, 9
    MyMESH2.Add_WallBack Vector(550, -1, 5500), Vector(5000, 10, 5500), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(500, 10, 5500), Vector(5000, 10, 5550), 18, 1, 9

    'add big wall on the south to limit city extension
    MyMESH2.Add_WallBack Vector(-5000, -1, 5551), Vector(5000, 1500, 5551), 18, 1, 10


   
    'then we build our mesh
    MyMESH2.BuildMesh

End Sub

Sub GameLoop()

 
    Do

        'Now we use camera
        GetInput
       
        'change the clear color randomely
        If QUEST.Get_KeyPressedApi(vbKeySpace) Then QUEST.Set_BackbufferClearColor (D3DColorXRGB(Rnd * 255, Rnd * 255, Rnd * 255))
        'we begin 3D
        QUEST.Begin3D

        MyMESH1.Render
        MyMESH2.Render

        'draw FPS
        QUEST.Draw_Text "FPS=" + CStr(QUEST.Get_FramesPerSeconde), 1, 10, &HFFFFFFFF

        If ShowInfo Then
            QUEST.Draw_Text "Polygon =" + CStr(QUEST.Get_NumberOfPolygonDrawn), 1, 25, &HFFFFFFFF
            QUEST.Draw_Text "Press ESC key to quit", 1, 40, &HFFFFFF00

            QUEST.Draw_Text "Press F1 to switch in FREE camera mode,F2 to FPS camera mode", 1, 55, &HFF00FF00

            QUEST.Draw_Text "Use Mouse to Rotate camera,Arrow Keyboard to move camera", 1, 85, &HFFFFFFFF

            If KAMERA.Get_CameraStyle = FPS_STYLE Then
                QUEST.Draw_Text "Current camera mode=", 1, 115, &HFFFFFFFF
                QUEST.Draw_Text "6 DEGREE OF FREEDOM STYLE", 170, 115, &HFFFF0001

              Else
                QUEST.Draw_Text "Current camera mode=", 1, 115, &HFFFFFFFF
                QUEST.Draw_Text "FIRST PERSON SHOOTER STYLE", 170, 115, &HFFFF0001

            End If

            QUEST.Draw_Text "Press 'H' to hide info,'S' to show info", 1, 135, &HFFFFFFFF

        End If

        'we close 3D Drawing
        QUEST.End3D
        DoEvents

        If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_ESCAPE) Then Call CloseGame
    Loop

End Sub

'in this sub we used
'advanced camera movements
Sub GetInput()

  'for informations printed to the screen

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_F1) Then KAMERA.Set_CameraStyle FREE_6DOF
    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_F2) Then KAMERA.Set_CameraStyle FPS_STYLE

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_S) Then ShowInfo = True
    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_H) Then ShowInfo = False

    'we use Get_TimePassed to move camera in X unit comparativelly to the
    'time passed, so are in Time based animation

    'strafe
    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_LEFT) Then _
       KAMERA.Strafe_Left QUEST.Get_TimePassed * PLAYER_SPEED

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_RIGHT) Then _
       KAMERA.Strafe_Right QUEST.Get_TimePassed * PLAYER_SPEED

    'move forward and backward
    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_UP) Then _
       KAMERA.Move_Forward QUEST.Get_TimePassed * PLAYER_SPEED

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_RCONTROL) Then _
       KAMERA.Move_Forward QUEST.Get_TimePassed * PLAYER_SPEED * 8

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_DOWN) Then _
       KAMERA.Move_Backward QUEST.Get_TimePassed * PLAYER_SPEED

    'here we use automated Camera rotation via Mouse Input
    '1st param=mouse speed default=0.001
    '2nd param=invert mouse default=false
    '3rd param=center mouse default=false
    'Rotate camera
    KAMERA.RotateByMouse 0.001, False, False

    KAMERA.Update

End Sub

'we quit game here
Sub CloseGame()

    MyMESH1.Free
    MyMESH2.Free
    QUEST.FreeEngine
    Set MyMESH1 = Nothing
    Set MyMESH2 = Nothing
    Set QUEST = Nothing
    Set KAMERA = Nothing
    Set KEY = Nothing
    

    End

End Sub

Private Sub Form_Unload(Cancel As Integer)

    CloseGame

End Sub
