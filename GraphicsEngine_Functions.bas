Attribute VB_Name = "GraphicsEngine_Functions"
Public Function InitGame()
    ReDim Lights(1 To 1) As Light
    NumLights = 1
    Lights(1).Color.R = 255
    Lights(1).Color.G = 255
    Lights(1).Color.B = 255
    Lights(1).Dir = Normalise(MakeVector(1, -0.5, 1))
    Lights(1).Pos = MakeVector(0, 4000, 0)
    Lights(1).Type = LIGHT_DIRECTION
    
    FrmMain.lLoading.Caption = "Loading...Sky"
    DoEvents
    
    Set SBox = New ESkyBox
    SBox.Init App.Path + "\Sky.bmp"
    
    Set Scene = New ELandScape
    
    InAir = False
    
    PHeight = PHei
End Function

Public Function LoadMission(MPath As String)
    Dim TT As String, FP As String, MP As String, MX As Single, MZ As Single, MMX As Single, MMZ As Single, Hei As Single
    NumObjects = 0
    Open MPath For Input As #2
        While EOF(2) = False
            Input #2, TT
            If UCase(TT) = "M" Then
                Input #2, MP, FP
                Scene.Init FP, MP
            ElseIf UCase(TT) = "O" Then
                FrmMain.lLoading.Caption = "Loading...Mission Object"
                DoEvents
                NumObjects = NumObjects + 1
                ReDim Preserve Objects(1 To NumObjects) As EObject
                Set Objects(NumObjects) = New EObject
                Input #2, MP
                Objects(NumObjects).Init MP
            ElseIf UCase(TT) = "W" Then
                NumWaters = NumWaters + 1
                ReDim Preserve Waters(1 To NumWaters) As EWater
                Set Waters(NumWaters) = New EWater
                Input #2, MX, MZ, MMX, MMZ, Hei
                Waters(NumWaters).Initialise MX, MZ, MMX, MMZ, Hei, App.Path + "\water.jpg"
            End If
        Wend
    Close #2
End Function

Public Function CheckObjectCollision(PStart As D3DVECTOR, PEnd As D3DVECTOR, POut As D3DVECTOR) As Boolean
    POut = PEnd
    If NumObjects > 0 Then
        For k = 1 To NumObjects
            If Objects(k).CheckCollision(PStart, PEnd, POut) = True Then
                CheckObjectCollision = True
            End If
        Next k
    End If
End Function

Public Function RenderGameItems()
    Scene.Render
    SBox.Render
    If NumObjects > 0 Then
        For k = 1 To NumObjects
            Objects(k).Render
        Next k
    End If
    If NumWeapons > 0 Then
        Weapons(CurrWeapon).Render
    End If
    If NumWaters > 0 Then
        For k = 1 To NumWaters
            Waters(k).Render
        Next k
    End If
End Function

Public Function RotateCamera(Alpha As Single, Beta As Single)
    Dim MatWorld2 As D3DMATRIX, matworld As D3DMATRIX
    Dim yy As Single
    D3DXMatrixRotationAxis MatWorld2, MakeVector(EyeDir.Z, 0, -EyeDir.X), Sin(Beta)
    D3DXMatrixRotationY matworld, Sin(Alpha)
    D3DXMatrixMultiply matworld, matworld, MatWorld2
    With EyeDir
        yy = .X * matworld.m12 + .Y * matworld.m22 + .Z * matworld.m32
        If yy < -0.9 Then
            yy = -0.9
        ElseIf yy > 0.9 Then
            yy = 0.9
        End If
        EyeDir = MakeVector(.X * matworld.m11 + .Y * matworld.m21 + .Z * matworld.m31, _
                                yy, _
                                .X * matworld.m13 + .Y * matworld.m23 + .Z * matworld.m33)
        EyeDir = Normalise(EyeDir)
    End With
End Function

Public Function Move(ForW As Single, Side As Single)
    Dim TempP As D3DVECTOR
    Dim Res As D3DVECTOR
    TempP = EyePos
    If InWater = False Then
        If ForW <> 0 Then
            TempP = MakeVector(EyePos.X + EyeDir.X * MoveSpeed * ForW, EyePos.Y, EyePos.Z + EyeDir.Z * MoveSpeed * ForW)
        End If
        If Side <> 0 Then
            TempP = MakeVector(EyePos.X + EyeDir.Z * MoveSpeed * Side, EyePos.Y, EyePos.Z - EyeDir.X * MoveSpeed * Side)
        End If
        CheckObjectCollision EyePos, TempP, TempP
        If Scene.RayCollisionDown(TempP, TempP) = True Then
            If TempP.Y < EyePos.Y - Gravity Then
                FallTime = FallTime + 1
                TempP.Y = EyePos.Y - Gravity * FallTime / 20
                OnFloor = False
            Else
                FallTime = 0
                OnFloor = True
            End If
            If InAir > 0 Then
                TempP.Y = TempP.Y + Gravity + JSpeed
                InAir = InAir - JSpeed
            End If
        Else
            TempP.Y = EyePos.Y - Gravity
        End If
        EyePos = TempP
    Else
        If ForW <> 0 Then
            TempP = MakeVector(EyePos.X + EyeDir.X * MoveSpeed * ForW, EyePos.Y + EyeDir.Y * ForW, EyePos.Z + EyeDir.Z * MoveSpeed * ForW)
        End If
        If Side <> 0 Then
            TempP = MakeVector(EyePos.X + EyeDir.Z * MoveSpeed * Side, EyePos.Y, EyePos.Z - EyeDir.X * MoveSpeed * Side)
        End If
        CheckObjectCollision EyePos, TempP, TempP
        EyePos = TempP
        If Scene.RayCollisionDown(EyePos, TempP) = True Then
            If TempP.Y >= EyePos.Y Then
                EyePos.Y = TempP.Y
            End If
        End If
    End If
End Function

Public Function MainLoop()
    Do While Finish = False
        CheckKey
        LastStandX = StandX
        LastStandY = StandY
        Move MoveF, MoveS
        Scene.UpdateLS StandX - LastStandX, StandY - LastStandY
        SBox.Update
        RenderAll
        DoEvents
    Loop
End Function
