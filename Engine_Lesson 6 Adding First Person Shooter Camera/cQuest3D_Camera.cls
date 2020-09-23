VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuest3D_Camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'
'       In this class module we define all variables needed for our Camera class
'
' THIS MODULE CONTAINS VERY USEFUL AND HARD HAND CODED METHODS, IF YOU
'   WANT TO USE THEM, GIVE CREDITS TO MY PERSON (Polaris),johna_pop@yahoo.fr
'   I HAVE SPENT A LOT OF TIME ON THAT CAMERA CLASS, IT IS THE BEST I HAVE CODED
'   AND THE BEST AVAILABLE IN VISUAL BASIC
'
'  USE IT, MODIFY IT BUT GIVE ME CREDIT
'
' Polaris: http://perso.numericable.fr/~gazkole/malakoff/index.htm
'
'================================================================================================

Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Dim g_fAngleX  As Single '  0.0     'Rotation angle for the x-axis
Dim g_fAngleY  As Single '  0.0     'Rotation angle for the y-axis
Dim g_fAngleZ  As Single '  0.0     'Rotation angle for the y-axis

Dim g_fPosX    As Single '  0.0     'Camera position on the x-axis
Dim g_fPosY    As Single '  1.2     'Camera position on the y-axis
Dim g_fPosZ    As Single ' -8.0     'Camera position on the z-axis

'for save datas
Dim Save_fAngleX  As Single '  0.0     'Rotation angle for the x-axis
Dim Save_fAngleY  As Single '  0.0     'Rotation angle for the y-axis
Dim Save_fAngleZ  As Single '  0.0     'Rotation angle for the y-axis

Dim Save_fPosX    As Single '  0.0     'Camera position on the x-axis
Dim Save_fPosY    As Single '  1.2     'Camera position on the y-axis
Dim Save_fPosZ    As Single ' -8.0     'Camera position on the z-axis

Private RotMat As D3DMATRIX
'Private RotQuat As D3DQUATERNION
'
'Private Save_RotMat As D3DMATRIX
'Private Save_RotQuat As D3DQUATERNION

Dim m_bMovedSinceLastUpdate As Boolean

Enum NEMO_PROJECTION_TYPE
    PT_ORTHOGRAPHIC
    PT_PERSPECTIVE_LH
    'PT_PERSPECTIVE_RH

End Enum

Enum NEMO_CAMERA_STYLE
    FREE_6DOF = 0
    FPS_STYLE = 1
End Enum

Dim CAM_TYPE As NEMO_CAMERA_STYLE

Dim m_ProjectionType As NEMO_PROJECTION_TYPE

Private M_vlastPOS As D3DVECTOR
Private M_vDir As D3DVECTOR
Private M_vVel As D3DVECTOR

Private Save_vlastPOS As D3DVECTOR
Private Save_vDir As D3DVECTOR
Private Save_vVel As D3DVECTOR

Private SAVEViewFrust As QUEST3D_FOV

Private Sub Class_Initialize()

    g_fAngleX = 0#       ' Rotation angle for the x-axis
    g_fAngleY = 0#       ' Rotation angle for the y-axis
    g_fAngleZ = 0#       ' Rotation angle for the y-axis

    g_fPosX = 0#         ' Camera position on the x-axis
    g_fPosY = 1.2        ' Camera position on the y-axis
    g_fPosZ = -8#        ' Camera position on the z-axis

    CAM_TYPE = FPS_STYLE

    m_ProjectionType = PT_PERSPECTIVE_LH

    Set_Position Vector(0, 0, 0)
    Set_LookAt Vector(0, 0, 1)

End Sub

Function Get_Direction() As D3DVECTOR

    Get_Direction = M_vDir 'Vector(Data.matView.m13, Data.matView.m23, Data.matView.m33) '

End Function

Function Get_LastPosition() As D3DVECTOR

    Get_LastPosition = M_vlastPOS

End Function

Sub RecallLastPosition()

    g_fPosX = M_vlastPOS.x
    g_fPosY = M_vlastPOS.y
    g_fPosZ = M_vlastPOS.z

    m_bMovedSinceLastUpdate = True

    Update

End Sub

Function Get_MatrixProjection() As D3DMATRIX

    Get_MatrixProjection = Data.MatProjec

End Function

Function Get_MatrixView() As D3DMATRIX

    Get_MatrixView = Data.matView

End Function

Function Get_Position() As D3DVECTOR

    Get_Position = Vector(g_fPosX, g_fPosY, g_fPosZ)

End Function

Function Get_ScreenHeight() As Long

    Get_ScreenHeight = Data.Buffer_Height

End Function

Function Get_ScreenWidth() As Long

    Get_ScreenWidth = Data.Buffer_Width

End Function

Sub Get_RayPick(ByVal x As Single, ByVal y As Single, ByRef Vorg As D3DVECTOR, Vdir As D3DVECTOR)

  Dim dx, dy
  Dim P1 As D3DVECTOR
  Dim P2 As D3DVECTOR

  Dim invMatrix As D3DMATRIX, ViewMatrix As D3DMATRIX

    dx = Tan(Data.ViewFrust.FovAngle * 0.5) * (x / (Data.Buffer_Width * 0.5) - 1#) / Data.ViewFrust.Aspect

    dy = Tan(Data.ViewFrust.FovAngle * 0.5) * (1# - y / (Data.Buffer_Height * 0.5))

    Call obj_Device.GetTransform(D3DTS_VIEW, ViewMatrix)

    Call D3DXMatrixInverse(invMatrix, 0, ViewMatrix)

    P1 = Vector(dx * Data.ViewFrust.Near, dy * Data.ViewFrust.Near, Data.ViewFrust.Near)
    P2 = Vector(dx * Data.ViewFrust.Far, dy * Data.ViewFrust.Far, Data.ViewFrust.Far)

    Call D3DXVec3TransformCoord(Vorg, P1, invMatrix)
    Call D3DXVec3TransformCoord(Vdir, P2, invMatrix)

End Sub

Function Get_Rotation2() As D3DQUATERNION

  Dim Q As D3DQUATERNION

    D3DXQuaternionRotationMatrix Q, RotMat

    Get_Rotation2.x = Q.x
    Get_Rotation2.y = Q.y
    Get_Rotation2.z = Q.z
    Get_Rotation2.w = Q.w

End Function

Function Get_RotationEuler() As D3DVECTOR

    Get_RotationEuler = Vector(g_fAngleX, g_fAngleY, g_fAngleZ)

End Function

Function Get_Velocity() As D3DVECTOR

  'Get_Velocity = M_vVel 'Vector(Data.matView.m13, Data.matView.m23, Data.matView.m33) '

    D3DXVec3Subtract Get_Velocity, Vector(g_fPosX, g_fPosY, g_fPosZ), Get_LastPosition

End Function

Function Get_VelocityNormalized() As D3DVECTOR

  Dim V As D3DVECTOR

    'Get_Velocity = M_vVel 'Vector(Data.matView.m13, Data.matView.m23, Data.matView.m33) '

    D3DXVec3Subtract V, Vector(g_fPosX, g_fPosY, g_fPosZ), Get_LastPosition
    D3DXVec3Normalize Get_VelocityNormalized, V

End Function

Function Get_ViewDirection() As D3DVECTOR

    Get_ViewDirection = Vector(Data.matView.m13, Data.matView.m23, Data.matView.m33) '

    'D3DXVec3Normalize Get_ViewDirection, Get_ViewDirection

End Function

Public Sub Move_Backward(ByVal StepAmount As Single)

    If CAM_TYPE = FREE_6DOF Then

        g_fPosX = g_fPosX + Sin((g_fAngleY)) * Cos(g_fAngleX) * StepAmount
        g_fPosY = g_fPosY - Sin(g_fAngleX) * StepAmount
        g_fPosZ = g_fPosZ - Cos((g_fAngleY)) * Cos(g_fAngleX) * StepAmount
      Else

        g_fPosX = g_fPosX + Sin((g_fAngleY)) * StepAmount
        g_fPosZ = g_fPosZ - Cos((g_fAngleY)) * StepAmount
    End If

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub Move_Forward(ByVal StepAmount As Single)

    If CAM_TYPE = FREE_6DOF Then

        g_fPosX = g_fPosX - Sin((g_fAngleY)) * Cos(g_fAngleX) * StepAmount
        g_fPosY = g_fPosY + Sin(g_fAngleX) * StepAmount
        g_fPosZ = g_fPosZ + Cos((g_fAngleY)) * Cos(g_fAngleX) * StepAmount
      Else

        g_fPosX = g_fPosX - Sin((g_fAngleY)) * StepAmount
        g_fPosZ = g_fPosZ + Cos((g_fAngleY)) * StepAmount
    End If

    m_bMovedSinceLastUpdate = True

End Sub

Sub Pop_CameraDatas()

    g_fAngleX = Save_fAngleX '= g_fAngleX 'Single '  0.0     'Rotation angle for the x-axis
    g_fAngleY = Save_fAngleY '= g_fAngleY 'Single '  0.0     'Rotation angle for the y-axis
    g_fAngleZ = Save_fAngleZ '= g_fAngleZ 'Single '  0.0     'Rotation angle for the y-axis

    g_fPosX = Save_fPosX '= g_fPosX 'Single '  0.0     'Camera position on the x-axis
    g_fPosY = Save_fPosY '= g_fPosY 'Single '  1.2     'Camera position on the y-axis
    g_fPosZ = Save_fPosZ '= g_fPosZ 'Single ' -8.0     'Camera position on the z-axis

    M_vlastPOS = Save_vlastPOS '= M_vlastPOS 'D3DVECTOR
    M_vDir = Save_vDir '= M_vDir 'D3DVECTOR
    M_vVel = Save_vVel '= M_vVel 'D3DVECTOR

    Data.ViewFrust = SAVEViewFrust

    m_bMovedSinceLastUpdate = 1

    Update

End Sub

Sub Push_CameraDatas()

    Save_fAngleX = g_fAngleX 'Single '  0.0     'Rotation angle for the x-axis
    Save_fAngleY = g_fAngleY 'Single '  0.0     'Rotation angle for the y-axis
    Save_fAngleZ = g_fAngleZ 'Single '  0.0     'Rotation angle for the y-axis

    Save_fPosX = g_fPosX 'Single '  0.0     'Camera position on the x-axis
    Save_fPosY = g_fPosY 'Single '  1.2     'Camera position on the y-axis
    Save_fPosZ = g_fPosZ 'Single ' -8.0     'Camera position on the z-axis

    Save_vlastPOS = M_vlastPOS 'D3DVECTOR
    Save_vDir = M_vDir 'D3DVECTOR
    Save_vVel = M_vVel 'D3DVECTOR

    SAVEViewFrust = Data.ViewFrust

End Sub

Public Sub Roll_Left(ByVal AngleAmount As Single)

    g_fAngleZ = g_fAngleZ + AngleAmount

    If (g_fAngleZ >= 360# * QUEST3D_RAD) Then

        g_fAngleZ = g_fAngleZ - 360# * QUEST3D_RAD
    End If

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub Roll_Right(ByVal AngleAmount As Single)

    g_fAngleZ = g_fAngleZ - AngleAmount

    If (g_fAngleZ < 0) Then

        g_fAngleZ = g_fAngleZ + 360# * QUEST3D_RAD
    End If

    m_bMovedSinceLastUpdate = True

End Sub

Sub Rotate(vRotationRadian As D3DVECTOR)

    g_fAngleX = vRotationRadian.x
    g_fAngleY = vRotationRadian.y
    g_fAngleZ = vRotationRadian.z

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub RotateByMouse(Optional ByVal MouseSpeed As Single = 0.001, Optional ByVal InvertMouse As Boolean = 0, Optional ByVal CenterMousePos As Boolean = 0)

    DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE

    ' Turn left or right
    If (DIMOUSESTATE.lX <> 0) Then

        g_fAngleY = g_fAngleY + (-DIMOUSESTATE.lX * MouseSpeed)

        If (g_fAngleY < 0#) Then

            g_fAngleY = 2 * QUEST3D_PI + g_fAngleY
          ElseIf (g_fAngleY >= 2 * QUEST3D_PI) Then

            g_fAngleY = g_fAngleY - 2 * QUEST3D_PI
        End If
    End If

    ' Look up or down
    If (DIMOUSESTATE.lY <> 0) Then

        If (InvertMouse) Then

            g_fAngleX = g_fAngleX + DIMOUSESTATE.lY * MouseSpeed

          Else

            g_fAngleX = g_fAngleX - DIMOUSESTATE.lY * MouseSpeed
        End If

        If (g_fAngleX < 0#) Then

            g_fAngleX = 2 * QUEST3D_PI + g_fAngleX

          ElseIf (g_fAngleX >= 3 * QUEST3D_PI) Then

            g_fAngleX = g_fAngleX - 2 * QUEST3D_PI
        End If

        If (g_fAngleX > QUEST3D_PI / 2# And g_fAngleX <= QUEST3D_PI) Then

            g_fAngleX = QUEST3D_PI / 2

          ElseIf (g_fAngleX > 180 * QUEST3D_RAD And g_fAngleX < 270# * QUEST3D_RAD) Then

            'g_fAngleX = 270# * QUEST3D_RAD
        End If
        
    End If

    m_bMovedSinceLastUpdate = True

    If CenterMousePos Then SetCursorPos (Data.Buffer_Rect.Right - Data.Buffer_Rect.Left) / 2 + Data.Buffer_Rect.Left, (Data.Buffer_Rect.bottom - Data.Buffer_Rect.Top) / 2 + Data.Buffer_Rect.Top

End Sub

Sub RotateX(ByVal DegreeAngle As Single)

    g_fAngleX = QUEST3D_RAD * DegreeAngle
    m_bMovedSinceLastUpdate = True

End Sub

Sub RotateY(ByVal DegreeAngle As Single)

    g_fAngleY = QUEST3D_RAD * DegreeAngle
    m_bMovedSinceLastUpdate = True

End Sub

Sub RotateZ(ByVal DegreeAngle As Single)

    g_fAngleZ = QUEST3D_RAD * DegreeAngle
    m_bMovedSinceLastUpdate = True

End Sub

Private Function ScreenTovec(ByVal Xp As Integer, ByVal Yp As Integer, ByVal m_fRadius As Single) As D3DVECTOR

  '// Scale to screen

  Dim x As Single, y As Single, sscale As Single, z As Single, mag As Single

    x = -(Xp - Data.Buffer_Width / 2) / (m_fRadius * Data.Buffer_Width / 2)
    y = (Yp - Data.Buffer_Height / 2) / (m_fRadius * Data.Buffer_Height / 2)

    z = 0#
    mag = x * x + y * y

    If (mag > 1#) Then

        sscale = 1# / Sqr(mag)
        x = x + x * sscale
        y = y + y * sscale

      Else
        z = Sqr(1# - mag)
    End If

    '// Return vector
    ScreenTovec = Vector(x, y, z)

End Function

'Private Function GetRotationFromCursor(fTrackBallRadius As Single) As D3DQUATERNION
'
'Dim sx As Single, sy As Single, sscale As Single, sZ As Single, Mag As Single
'Dim PT As POINTAPI, d1 As Single, d2 As Single, t As Single
'Dim Quat As D3DQUATERNION
'Dim P1 As D3DVECTOR
'Dim P2 As D3DVECTOR, Vaxis As D3DVECTOR, vecDiff As D3DVECTOR
'Dim Fangle As Single
'
'If Data.Buffer_Rect.Right = 0 Then Exit Function
'
'Quat.w = 1
'
''Data
'
'    Call GetCursorPos(PT)
'    'GetClientRect( hWnd, rc )
'    'Call ScreenToClient(Data.Hwindow, PT)
'    sx = (((2# * PT.X) / (Data.Buffer_Rect.Right - Data.Buffer_Rect.Left)) - 1)
'    sy = (((2# * PT.Y) / (Data.Buffer_Rect.Bottom - Data.Buffer_Rect.Top)) - 1)
'
'
'    If (sx = 0# And sy = 0#) Then
'      GetRotationFromCursor = Quat
'      Exit Function
'    End If
'
'    d2 = Sqr(sx * sx + sy * sy)
'
'    If (d2 < fTrackBallRadius * 0.707106781186548) Then _
    '        sZ = Sqr(fTrackBallRadius * fTrackBallRadius - d2 * d2) _
    '    Else _
    '        sZ = (fTrackBallRadius * fTrackBallRadius) / (2# * d2)
'
'    ' Get two points on trackball's sphere
'     P1 = Vector(sx, sy, sZ)
'     P2 = Vector(0#, 0#, fTrackBallRadius)
'
'    ' Get axis of rotation, which is cross product of p1 and p2
'
'    Call D3DXVec3Cross(Vaxis, P1, P2)
'
'    ' Calculate angle for the rotation about that axis
'     vecDiff.X = P2.X - P1.X
'     vecDiff.Y = P2.Y - P1.Y
'     vecDiff.Z = P2.Z - P1.Z
'
'
'    t = D3DXVec3Length(vecDiff) / (2# * fTrackBallRadius)
'    If (t > 1#) Then t = 1#
'    If (t < -1#) Then t = -1#
'    Fangle = 2# * farcSin(t)
'
'    ' Convert axis to quaternion
'    Quat.w = 0
'
'    Call D3DXQuaternionRotationAxis(Quat, Vaxis, Fangle)
'    GetRotationFromCursor = Quat
'
'End Function
'

Sub Set_camera(vPos As D3DVECTOR, LookAt As D3DVECTOR)

    If VECTOR_Compare(vPos, Vector(g_fPosX, g_fPosY, g_fPosZ), 0.005) Then Exit Sub

    g_fPosX = vPos.x
    g_fPosY = vPos.y
    g_fPosZ = vPos.z

    Set_LookAt LookAt

End Sub

Sub Set_CameraProjectionType(ProjType As NEMO_PROJECTION_TYPE)

    m_ProjectionType = ProjType
    UpdateFrustum

End Sub

Function Get_CameraProjectionType() As NEMO_PROJECTION_TYPE

    Get_CameraProjectionType = m_ProjectionType

End Function

Sub Set_CameraStyle(Style As NEMO_CAMERA_STYLE)

    CAM_TYPE = Style

End Sub

Function Get_CameraStyle() As NEMO_CAMERA_STYLE

    Get_CameraStyle = CAM_TYPE

End Function

Sub Set_Direction(Vdir As D3DVECTOR)

  Dim Vat As D3DVECTOR
  Dim Vlook As D3DVECTOR
  Dim VANG As D3DVECTOR

    Vlook = Get_Position
    Vat = Vector(Vlook.x + Vdir.x * 10, Vlook.y + Vdir.y * 10, Vlook.z + Vdir.z * 10)

    GetRotationFromTO Vlook, Vat, VANG

    g_fAngleX = VANG.x
    g_fAngleY = VANG.y
    g_fAngleZ = VANG.z

    m_bMovedSinceLastUpdate = True

End Sub

Sub Set_LookAt(Vlook As D3DVECTOR)

    m_bMovedSinceLastUpdate = True

    '  Dim VANG As D3DVECTOR
    '  Dim fXRot As Single, fYRot As Single, fXDiff As Single, fYDiff As Single, fZDiff As Single
  Dim currentVelocity As D3DVECTOR
    '      fXDiff = Vlook.x - g_fPosX
    '      fYDiff = Vlook.y - g_fPosY
    '      fZDiff = Vlook.z - g_fPosZ
    '
    '      fXDiff = g_fPosX - Vlook.x
    '      fYDiff = g_fPosY - Vlook.y
    '      fZDiff = g_fPosZ - Vlook.z

    'D3DXVec3Subtract currentVelocity, Vlook, Vector(g_fPosX, g_fPosY, g_fPosZ)

    currentVelocity.x = Vlook.x - g_fPosX
    currentVelocity.y = Vlook.y - g_fPosY
    currentVelocity.z = Vlook.z - g_fPosZ

    D3DXVec3Normalize currentVelocity, currentVelocity
    '    '
    '    '    'GetRotationFromTO Vector(g_fPosX, g_fPosY, g_fPosZ), Vlook, VANG
    '    '
    '        g_fAngleY = Arctan(-fYDiff, sqr(fXDiff * fXDiff + fZDiff * fZDiff))
    '        g_fAngleX = Arctan(fXDiff, fZDiff)
    '
    '        'g_fAngleX = VANG.x
    ''    g_fAngleY = VANG.y
    ''    g_fAngleZ = VANG.z

    '
    '        If (currentVelocity.x <> 0#) Then
    '
    '            g_fAngleY = -atn(currentVelocity.z / currentVelocity.x)
    '
    ''        ElseIf (currentVelocity.x < 0#) Then
    ''
    ''                 g_fAngleY = -atn(currentVelocity.z / currentVelocity.x)
    ''
    ''        ElseIf (currentVelocity.x = 0#) Then
    ''
    ''                g_fAngleY = 0
    '        End If
    '
    '
    '     If currentVelocity.x <> 0 Or currentVelocity.z <> 0 Then
    '        g_fAngleY = atn(currentVelocity.y / _
        '                sqr(currentVelocity.z * currentVelocity.z _
        '                + currentVelocity.x * currentVelocity.x))
    '    End If

    If currentVelocity.z <> 0 Or currentVelocity.x <> 0 Then
        g_fAngleX = Atn(currentVelocity.y / _
                    Sqr(currentVelocity.z * currentVelocity.z _
                    + currentVelocity.x * currentVelocity.x))
    End If

    If currentVelocity.z <> 0 Then
        g_fAngleY = -ArcTan(currentVelocity.x, currentVelocity.z)
    End If

End Sub

Sub Set_Position(vPos As D3DVECTOR)

    If VECTOR_Compare(vPos, Vector(g_fPosX, g_fPosY, g_fPosZ), 0.005) Then Exit Sub

    g_fPosX = vPos.x
    g_fPosY = vPos.y
    g_fPosZ = vPos.z

    m_bMovedSinceLastUpdate = True

End Sub

Sub Set_Position2(x As Single, y As Single, z As Single)

    If VECTOR_Compare(Vector(x, y, z), Vector(g_fPosX, g_fPosY, g_fPosZ), 0.005) Then Exit Sub

    g_fPosX = x
    g_fPosY = y
    g_fPosZ = z

    m_bMovedSinceLastUpdate = True

End Sub

Sub Set_ViewFrustum(ByVal Near As Single, ByVal Far As Single, ByVal FovAngle As Single, Optional ByVal Aspect As Single = -1)

    If Aspect = -1 Then Aspect = Data.Buffer_Width / Data.Buffer_Height
    With Data.ViewFrust
        .Aspect = Aspect
        .Near = Near
        .Far = Far
        .FovAngle = FovAngle

    End With

    Data.FRUSTUM_HASCHANGED = True
    'UpdateFrustum

End Sub

Function Get_FovAngle() As Single

    Get_FovAngle = Data.ViewFrust.FovAngle

End Function

Public Sub Strafe_Down(ByVal StepAmount As Single)

    g_fPosY = g_fPosY - StepAmount

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub Strafe_Left(ByVal StepAmount As Single)

    g_fPosX = g_fPosX - Sin((g_fAngleY + 90 * QUEST3D_RAD#)) * StepAmount
    g_fPosZ = g_fPosZ + Cos((g_fAngleY + 90 * QUEST3D_RAD)) * StepAmount

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub Strafe_Right(ByVal StepAmount As Single)

    g_fPosX = g_fPosX - Sin((g_fAngleY - 90 * QUEST3D_RAD)) * StepAmount
    g_fPosZ = g_fPosZ + Cos((g_fAngleY - 90 * QUEST3D_RAD)) * StepAmount

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub Strafe_Up(ByVal StepAmount As Single)

    g_fPosY = g_fPosY + StepAmount

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub Turn_Down(ByVal AngleAmount As Single)

    g_fAngleX = g_fAngleX - AngleAmount

    If (g_fAngleX < 0) Then

        g_fAngleX = g_fAngleX + 360# * QUEST3D_RAD
    End If

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub Turn_Left(ByVal AngleAmount As Single)

    g_fAngleY = g_fAngleY + AngleAmount

    If (g_fAngleY >= 360# * QUEST3D_RAD) Then

        g_fAngleY = g_fAngleY - 360# * QUEST3D_RAD
    End If

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub Turn_Right(ByVal AngleAmount As Single)

    g_fAngleY = g_fAngleY - AngleAmount

    If (g_fAngleY < 0) Then

        g_fAngleY = g_fAngleY + 360# * QUEST3D_RAD
    End If

    m_bMovedSinceLastUpdate = True

End Sub

Public Sub Turn_Up(ByVal AngleAmount As Single)

    g_fAngleX = g_fAngleX + AngleAmount

    If (g_fAngleX >= 360# * QUEST3D_RAD) Then

        g_fAngleX = g_fAngleX - 360# * QUEST3D_RAD
    End If

    m_bMovedSinceLastUpdate = True

End Sub

Sub Update(Optional ByVal UpdateViewFrustum As Boolean = True)

    If Not m_bMovedSinceLastUpdate Then Exit Sub

  Dim vecEyePoint As D3DVECTOR     'The eye point
  Dim vecLookatPoint As D3DVECTOR  'The camera look-at target
  Dim vecUp As D3DVECTOR           'The current world's up
  Dim matRotationX As D3DMATRIX     'The matrix that rotates around the x-axis
  Dim matRotationY As D3DMATRIX     'The matrix that rotates around the y-axis
  Dim matRotationZ As D3DMATRIX     'The matrix that rotates around the y-axis

    'Dim matRotation As D3DMATRIX      'The matrix that rotates around the x and y-axis
  Dim matWorld As D3DMATRIX         'The world transformation matrix
  Dim matView As D3DMATRIX          'The view transformation matrix

    If Not VECTOR_Compare(M_vlastPOS, Vector(g_fPosX, g_fPosY, g_fPosZ), QUEST3D_EPSILON) Then

        M_vlastPOS = Data.EYES.EYE 'Vector(g_fPosX, g_fPosY, g_fPosZ)
        M_vDir = Vector(g_fPosX - M_vlastPOS.x, g_fPosY - M_vlastPOS.y, g_fPosZ - M_vlastPOS.z)

        Data.EYES.Dir = M_vDir
        M_vVel = M_vDir

        D3DXVec3Normalize M_vDir, M_vDir

        M_vVel = M_vDir
        'D3DXVec3Normalize M_vDir, M_vDir
        Data.EYES.DirNormalized = M_vDir
    End If

    vecEyePoint = Vector(g_fPosX, g_fPosY, g_fPosZ)
    vecLookatPoint = Vector(g_fPosX, g_fPosY, g_fPosZ + 0.001)
    vecUp = Vector(0#, 1#, 0#)

    'update datas
    Data.EYES.EYE = vecEyePoint
    Data.EYES.Dest_at = vecLookatPoint

    Call D3DXMatrixLookAtLH(matView, vecEyePoint, vecLookatPoint, vecUp)

    '    D3DXMatrixIdentity matView
    '    D3DXMatrixTranslation matView, g_fPosX, g_fPosY, g_fPosZ
    '

    Call D3DXMatrixRotationX(matRotationX, (g_fAngleX))
    Call D3DXMatrixRotationY(matRotationY, (g_fAngleY))
    Call D3DXMatrixRotationZ(matRotationZ, (g_fAngleZ))

    Data.EYES.AngX = g_fAngleX
    Data.EYES.ANGy = g_fAngleY
    Data.EYES.ANGz = g_fAngleZ

    Call D3DXMatrixMultiply(RotMat, matRotationY, matRotationX)
    Call D3DXMatrixMultiply(RotMat, RotMat, matRotationZ)

    Call D3DXMatrixMultiply(matView, matView, RotMat)

    Call obj_Device.SetTransform(D3DTS_VIEW, matView)
    Data.matView = matView

    m_bMovedSinceLastUpdate = False

    If Data.FRUSTUM_HASCHANGED Then UpdateFrustum

    If UpdateViewFrustum Then _
       SetUpFrustum

End Sub

Private Sub UpdateFrustum()

    If m_ProjectionType = PT_PERSPECTIVE_LH Then
        D3DXMatrixPerspectiveFovLH Data.MatProjec, Data.ViewFrust.FovAngle, Data.ViewFrust.Aspect, Data.ViewFrust.Near, Data.ViewFrust.Far

        'ElseIf m_ProjectionType = PT_PERSPECTIVE_RH Then
        ' D3DXMatrixPerspectiveFovRH Data.MatProjec, Data.ViewFrust.FovAngle, Data.ViewFrust.Aspect, Data.ViewFrust.Near, Data.ViewFrust.Far

      Else
        Call D3DXMatrixOrthoLH(Data.MatProjec, Data.Buffer_Width, Data.Buffer_Height, 0#, Data.ViewFrust.Far)
        'Call D3DXMatrixOrthoLH(Data.MatProjec, 0#, Data.Buffer_Width, 0#, Data.Buffer_Height, 0#, 1#)

    End If

    'm_matProj = Data.MatProjec

    obj_Device.SetTransform D3DTS_PROJECTION, Data.MatProjec

    SetUpFrustum
    Data.FRUSTUM_HASCHANGED = 0

End Sub
