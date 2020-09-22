VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GWeapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Damage As Single
Private ReloadTime As Single
Private CoolDownTime As Single
Private LastFired As Single

Private Tris() As ATriangle
Private NumTris As Integer
Private Textures() As Direct3DTexture8
Private NumTex As Integer
Private Mouth As D3DVECTOR
Private FTris(1 To 5) As ATriangle

Private MaxBullets As Long
Private MaxClips As Long

Public Bullets As Long
Public Clips As Long

Private Vis As Boolean
Private Available As Boolean

Private ReCoil As Single
Private Fired As Boolean

Private FireSound As ESound
Private ReloadSound As ESound

Private WName As String
Private Reloading As Long

Private Const ReC As Single = 10
Private Const ReL As Single = 2

Public Function Init(WPath As String, sName As String, MaxB As Long, MaxC As Long)
    WName = sName
    MaxBullets = MaxB
    Reloading = 0
    MaxClips = MaxC
    Vis = False
    Dim TT As String, CT As ATriangle, TTex As String, Col As Long, Sho As String, Re As String
    Open App.Path + WPath For Input As #5
        While EOF(5) = False
            Input #5, TT
            If UCase(TT) = "T" Then
                NumTris = NumTris + 1
                ReDim Preserve Tris(1 To NumTris) As ATriangle
                With Tris(NumTris)
                    Input #5, .Vecs(1).Position.X, .Vecs(1).Position.Y, .Vecs(1).Position.Z, _
                              .Vecs(2).Position.X, .Vecs(2).Position.Y, .Vecs(2).Position.Z, _
                              .Vecs(3).Position.X, .Vecs(3).Position.Y, .Vecs(3).Position.Z, _
                              .Vecs(1).tu, .Vecs(1).tv, .Vecs(2).tu, .Vecs(2).tv, .Vecs(3).tu, .Vecs(3).tv, _
                              .Tex
                End With
                Set Tris(NumTris).mVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, 0, D3DFVF_VERTEX, D3DPOOL_DEFAULT)
            ElseIf UCase(TT) = "M" Then
                Input #5, TTex
                NumTex = NumTex + 1
                ReDim Preserve Textures(1 To NumTex) As Direct3DTexture8
                Set Textures(NumTex) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + TTex)
            ElseIf UCase(TT) = "S" Then
                Input #5, Sho, Re, Col
                Set FireSound = New ESound
                Set ReloadSound = New ESound
                FireSound.Create App.Path + Sho
                ReloadSound.Create App.Path + Re
                CoolDownTime = Col
            ElseIf UCase(TT) = "O" Then
                Input #5, Mouth.X, Mouth.Y, Mouth.Z
            End If
        Wend
    Close #5
    With FTris(1)
        .Vecs(1) = MakeVertex(Mouth.X, 10 + Mouth.Y, 3 + Mouth.Z, 0.5, 0)
        .Vecs(2) = MakeVertex(-2 + Mouth.X, Mouth.Y, Mouth.Z, 0, 1)
        .Vecs(3) = MakeVertex(2 + Mouth.X, Mouth.Y, Mouth.Z, 1, 1)
    End With
    With FTris(2)
        .Vecs(1) = MakeVertex(Mouth.X, -10 + Mouth.Y, 3 + Mouth.Z, 0.5, 0)
        .Vecs(2) = MakeVertex(-2 + Mouth.X, Mouth.Y, Mouth.Z, 0, 1)
        .Vecs(3) = MakeVertex(2 + Mouth.X, Mouth.Y, Mouth.Z, 1, 1)
    End With
    With FTris(3)
        .Vecs(1) = MakeVertex(-10 + Mouth.X, Mouth.Y, 3 + Mouth.Z, 0.5, 0)
        .Vecs(2) = MakeVertex(Mouth.X, 2 + Mouth.Y, Mouth.Z, 1, 1)
        .Vecs(3) = MakeVertex(Mouth.X, -2 + Mouth.Y, Mouth.Z, 0, 1)
    End With
    With FTris(4)
        .Vecs(1) = MakeVertex(10 + Mouth.X, Mouth.Y, 3 + Mouth.Z, 0.5, 0)
        .Vecs(2) = MakeVertex(Mouth.X, 2 + Mouth.Y, Mouth.Z, 1, 1)
        .Vecs(3) = MakeVertex(Mouth.X, -2 + Mouth.Y, Mouth.Z, 0, 1)
    End With
    With FTris(5)
        .Vecs(1) = MakeVertex(Mouth.X, Mouth.Y, Mouth.Z + 50000, 0.5, 0)
        .Vecs(2) = MakeVertex(Mouth.X - 1, Mouth.Y, Mouth.Z, 0, 1)
        .Vecs(3) = MakeVertex(Mouth.X + 1, Mouth.Y, Mouth.Z, 1, 1)
    End With
    NumTex = NumTex + 1
    ReDim Preserve Textures(1 To NumTex) As Direct3DTexture8
    Set Textures(NumTex) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\weapons\blast.mtx")
    For k = 1 To 5
        FTris(k).Tex = NumTex
        Set FTris(k).mVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, 0, D3DFVF_VERTEX, D3DPOOL_DEFAULT)
        For t = 1 To 3
            With FTris(k)
                .Vecs(t).Color = &HFFFFFFFF
                .vVecs(t) = .Vecs(t)
            End With
        Next t
    Next k
    For k = 1 To NumTris
        For t = 1 To 3
            With Tris(k)
                .vVecs(t) = .Vecs(t)
            End With
        Next t
    Next k
End Function

Private Function RefreshAll()
    Dim Dif As D3DVECTOR, Col As Long, Leng As Single
    Dim MatWorld2 As D3DMATRIX, matworld As D3DMATRIX
    If ReCoil > 0 And Timer - Reloading <= ReL Then
        ReCoil = ReCoil - (Timer - Reloading) / ReL
    End If
    If ReCoil > 0 And CoolDownTime > 0 Then
        ReCoil = ReCoil - (Timer - LastFired) / CoolDownTime
    Else
        ReCoil = 0
    End If
    Dif = Normalise(MakeVector(EyeDir.X, 0, EyeDir.Z))
    Leng = Acos(Dif.Z)
    If Dif.X < 0 Then Leng = 2 * Pi - Leng
    D3DXMatrixRotationAxis MatWorld2, MakeVector(EyeDir.Z, 0, -EyeDir.X), -EyeDir.Y
    D3DXMatrixRotationY matworld, Leng
    D3DXMatrixMultiply matworld, matworld, MatWorld2
    Col = Scene.GetTCol(EyePos)
    If InWater = True Then
        Col = GetCol(Col).R / 3
        Col = RGB(WatCol, Col, Col)
    End If
    For k = 1 To NumTris
        For t = 1 To 3
            With Tris(k).Vecs(t).Position
                Tris(k).vVecs(t).Position.X = .X * matworld.m11 + .Y * matworld.m21 + (.Z - ReCoil) * matworld.m31
                Tris(k).vVecs(t).Position.Y = .X * matworld.m12 + .Y * matworld.m22 + (.Z - ReCoil) * matworld.m32
                Tris(k).vVecs(t).Position.Z = .X * matworld.m13 + .Y * matworld.m23 + (.Z - ReCoil) * matworld.m33
                Tris(k).vVecs(t).Position.X = Tris(k).vVecs(t).Position.X + EyePos.X
                Tris(k).vVecs(t).Position.Y = Tris(k).vVecs(t).Position.Y + EyePos.Y + PHeight
                Tris(k).vVecs(t).Position.Z = Tris(k).vVecs(t).Position.Z + EyePos.Z
            End With
            Tris(k).vVecs(t).Color = Col
        Next t
    Next k
    For k = 1 To 5
        For t = 1 To 3
            With FTris(k).Vecs(t).Position
                FTris(k).vVecs(t).Position.X = .X * matworld.m11 + .Y * matworld.m21 + .Z * matworld.m31
                FTris(k).vVecs(t).Position.Y = .X * matworld.m12 + .Y * matworld.m22 + .Z * matworld.m32
                FTris(k).vVecs(t).Position.Z = .X * matworld.m13 + .Y * matworld.m23 + .Z * matworld.m33
                FTris(k).vVecs(t).Position.X = FTris(k).vVecs(t).Position.X + EyePos.X
                FTris(k).vVecs(t).Position.Y = FTris(k).vVecs(t).Position.Y + EyePos.Y + PHeight
                FTris(k).vVecs(t).Position.Z = FTris(k).vVecs(t).Position.Z + EyePos.Z
            End With
        Next t
    Next k
    ReloadVB
End Function

Public Function Render()
    If Vis = True Then
        If NumTris > 0 Then
            RefreshAll
            For k = 1 To NumTris
                D3DDevice.SetStreamSource 0, Tris(k).mVB, VertexSizeInBytes
                D3DDevice.SetVertexShader D3DFVF_VERTEX
                D3DDevice.SetTexture 0, Textures(Tris(k).Tex)
                D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 1
            Next k
        End If
        If Fired = True Then
            For k = 1 To 5
                D3DDevice.SetStreamSource 0, FTris(k).mVB, VertexSizeInBytes
                D3DDevice.SetVertexShader D3DFVF_VERTEX
                D3DDevice.SetTexture 0, Textures(NumTex)
                D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 1
                D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
            Next k
            Fired = False
        End If
    End If
    Dim matworld As D3DMATRIX
    D3DXMatrixIdentity matworld
    D3DDevice.SetTransform D3DTS_WORLD, matworld
End Function

Public Function UseThis() As Boolean
    If Available = True Then
        Vis = True
    End If
    UseThis = Available
    UpdateIBar
End Function

Public Function StashAway()
    If Available = True Then
        Vis = False
    End If
    UpdateIBar
End Function

Public Function Reload()
    If Clips > 0 Then
        Bullets = MaxBullets
        Clips = Clips - 1
        LastFired = Timer
        ReloadSound.PlaySound False
        Reloading = Timer
        ReCoil = ReC
    End If
    UpdateIBar
End Function

Public Function MakeAvailable()
    Available = True
End Function

Public Function Shoot()
    If Bullets > 0 And InWater = False Then
        If Timer - LastFired >= CoolDownTime Then
            LastFired = Timer
            FireSound.PlaySound False
            Bullets = Bullets - 1
            ReCoil = ReC
            Fired = True
        End If
    End If
    UpdateIBar
End Function

Private Function ReloadVB()
    For k = 1 To NumTris
        D3DVertexBuffer8SetData Tris(k).mVB, 0, VertexSizeInBytes * 3, 0, Tris(k).vVecs(1)
    Next k
    For k = 1 To 5
        D3DVertexBuffer8SetData FTris(k).mVB, 0, VertexSizeInBytes * 3, 0, FTris(k).vVecs(1)
    Next k
End Function
