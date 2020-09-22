VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ELandScape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private TileSize As Single
Private Tiles() As Tile
Private NumTiles As Integer
Private Texture As Direct3DTexture8
Private NumVertices As Double
Private Vertices() As VERTEX
Private TVerts() As VERTEX
Private mVB As Direct3DVertexBuffer8

Public Function GetTCol(P As D3DVECTOR) As Long
    Dim Temp As Single
    Dim SX As Long, SY As Long
    Dim StartP As D3DVECTOR
    StartP = P
    Temp = StartP.X / TileSize
    SX = CLng(Temp)
    If Temp > SX Then SX = SX + 1
    If SX <= 0 Then
        While SX <= 0
            SX = SX + NumTiles
            StartP.X = StartP.X + NumTiles * TileSize
        Wend
    ElseIf SX > NumTiles Then
        While SX > NumTiles
            SX = SX - NumTiles
            StartP.X = StartP.X - NumTiles * TileSize
        Wend
    End If
    Temp = StartP.Z / TileSize
    SY = CLng(Temp)
    If Temp > SY Then SY = SY + 1
    If SY <= 0 Then
        While SY <= 0
            SY = SY + NumTiles
            StartP.Z = StartP.Z + NumTiles * TileSize
        Wend
    ElseIf SY > NumTiles Then
        While SY > NumTiles
            SY = SY - NumTiles
            StartP.Z = StartP.Z - NumTiles * TileSize
        Wend
    End If
    GetTCol = Tiles(SX, SY).Vecs(1).Color
End Function

Public Function Init(FPath As String, MPath As String)
    Set Texture = D3DX.CreateTextureFromFile(D3DDevice, App.Path + FPath)
    
    Randomize Timer
    
    Dim Map(0 To 40, 0 To 40) As Single, MM As Single
    
    LoadHMap Map, App.Path + MPath
    
    LoadLandscape Map, 39, 400
    
    If FrmMain.cShadows.Value = 1 Then
        FrmMain.lLoading.Caption = "Loading...Shadows"
        DoEvents
        Lighting True
    End If
    
    ReloadWholeVB
End Function

Public Function LoadHMap(Arr() As Single, FPath As String)
    With FrmMain.HMap
        .Picture = LoadPicture(FPath)
        For k = 0 To 40
            For t = 0 To 40
                Arr(k, t) = GetCol(.Point(k, t)).R * 6
            Next t
        Next k
        .Cls
    End With
End Function

Public Function UpdateLS(DirX As Single, DirY As Single)
Dim CV As Long, CCV As Long
If DirX <> 0 Then
    If DirX < 0 Then
        For k = 1 To NumTiles
            For t = 1 To 6
                CCV = ((NumTiles - 1) * NumTiles)
                TVerts((k - 1) * 6 + t - 1) = Vertices(CCV * 6 + ((k - 1) * 6) + t - 1)
            Next t
        Next k
        For k = NumTiles - 1 To 1 Step -1
            For t = NumTiles To 1 Step d - 1
                For ps = 6 To 1 Step -1
                    CCV = k * NumTiles
                    Vertices((CCV * 6) + ((t - 1) * 6) + ps - 1) = Vertices(((CCV - NumTiles) * 6) + ((t - 1) * 6) + ps - 1)
                Next ps
            Next t
        Next k
        For k = 1 To NumTiles
            For t = 1 To 6
                CV = ((k - 1) * 6) + t - 1
                Vertices(CV) = TVerts(CV)
                Vertices(CV).Position.X = Vertices(CV).Position.X - (NumTiles * TileSize)
            Next t
        Next k
    Else
        For k = 1 To NumTiles
            For t = 1 To 6
                TVerts((k - 1) * 6 + t - 1) = Vertices(((k - 1) * 6) + t - 1)
            Next t
        Next k
        For k = 1 To NumTiles - 1
            For t = 1 To NumTiles
                For ps = 1 To 6
                    CCV = k * NumTiles
                    Vertices(((CCV - NumTiles) * 6) + ((t - 1) * 6) + ps - 1) = Vertices((CCV * 6) + ((t - 1) * 6) + ps - 1)
                Next ps
            Next t
        Next k
        For k = 1 To NumTiles
            For t = 1 To 6
                CCV = NumTiles * 6
                CV = ((NumTiles - 1) * CCV) + ((k - 1) * 6) + t - 1
                Vertices(CV) = TVerts(((k - 1) * 6) + t - 1)
                Vertices(CV).Position.X = Vertices(CV).Position.X + (NumTiles * TileSize)
            Next t
        Next k
    End If
End If

If DirY <> 0 Then
    If DirY < 0 Then
        For k = 1 To NumTiles
            For t = 1 To 6
                CCV = k * NumTiles
                TVerts((k - 1) * 6 + t - 1) = Vertices(((CCV - NumTiles) * 6) + ((NumTiles - 1) * 6) + t - 1)
            Next t
        Next k
        For t = NumTiles To 1 Step -1
            For k = NumTiles - 1 To 1 Step -1
                For ps = 6 To 1 Step -1
                    CCV = t * NumTiles
                    Vertices(((CCV - NumTiles) * 6) + (k * 6) + ps - 1) = Vertices(((CCV - NumTiles) * 6) + ((k - 1) * 6) + ps - 1)
                Next ps
            Next k
        Next t
        For k = 1 To NumTiles
            For t = 1 To 6
                CCV = k * NumTiles
                CV = ((CCV - NumTiles) * 6) + t - 1
                Vertices(CV) = TVerts(((k - 1) * 6) + t - 1)
                Vertices(CV).Position.Z = Vertices(CV).Position.Z - (NumTiles * TileSize)
            Next t
        Next k
    Else
        For k = 1 To NumTiles
            For t = 1 To 6
                CCV = k * NumTiles
                TVerts((k - 1) * 6 + t - 1) = Vertices(((CCV - NumTiles) * 6) + t - 1)
            Next t
        Next k
        For t = 1 To NumTiles
            For k = 1 To NumTiles - 1
                For ps = 1 To 6
                    CCV = t * NumTiles
                    Vertices(((CCV - NumTiles) * 6) + ((k - 1) * 6) + ps - 1) = Vertices(((CCV - NumTiles) * 6) + (k * 6) + ps - 1)
                Next ps
            Next k
        Next t
        For k = 1 To NumTiles
            For t = 1 To 6
                CCV = k * NumTiles
                CV = ((CCV - NumTiles) * 6) + ((NumTiles - 1) * 6) + t - 1
                Vertices(CV) = TVerts(((k - 1) * 6) + t - 1)
                Vertices(CV).Position.Z = Vertices(CV).Position.Z + (NumTiles * TileSize)
            Next t
        Next k
    End If
End If
ReloadVB
End Function

Private Function ReloadTile(DX As Long, DY As Long)
    With Tiles(DX, DY)
        For k = 1 To 6
            Vertices(((DX - 1) * NumTiles * 6) + ((DY - 1) * 6) + k - 1) = .Vecs(k)
        Next k
    End With
End Function

Private Function ReloadWholeVB()
    NumVertices = NumTiles * NumTiles
    NumVertices = (NumVertices * 6) + 2
    ReDim Vertices(0 To NumVertices) As VERTEX

    ReDim TVerts(0 To NumTiles * 6 - 1) As VERTEX

    Set mVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * NumVertices, 0, D3DFVF_VERTEX, D3DPOOL_DEFAULT)

    For k = 1 To NumTiles
        For t = 1 To NumTiles
            For P = 1 To 6
                With Tiles(k, t)
                    Vertices(((k - 1) * NumTiles * 6) + ((t - 1) * 6) + P - 1) = .Vecs(P)
                End With
            Next P
        Next t
    Next k
    
    D3DVertexBuffer8SetData mVB, 0, VertexSizeInBytes * NumVertices, 0, Vertices(0)
End Function

Private Function ReloadVB()
    D3DVertexBuffer8SetData mVB, 0, VertexSizeInBytes * NumVertices, 0, Vertices(0)
End Function

Private Function LoadLandscape(HMap() As Single, MaxTiles As Long, Optional Tilesiz As Single = 10)
    Dim km As Long, tm As Long
    Dim Shade As Single
    TileSize = Tilesiz
    NumTiles = MaxTiles
    ReDim Tiles(1 To NumTiles, 1 To NumTiles) As Tile
    For k = 1 To NumTiles
        km = k - 1
        For t = 1 To NumTiles
            With Tiles(k, t)
                tm = t - 1
                .Vecs(1) = MakeVertex(km * TileSize, HMap(km, tm), tm * TileSize, km / NumTiles, tm / NumTiles)
                .Vecs(4) = .Vecs(1)
                .Vecs(2) = MakeVertex(k * TileSize, HMap(k, tm), tm * TileSize, k / NumTiles, tm / NumTiles)
                .Vecs(3) = MakeVertex(k * TileSize, HMap(k, t), t * TileSize, k / NumTiles, t / NumTiles)
                .Vecs(5) = .Vecs(3)
                .Vecs(6) = MakeVertex(km * TileSize, HMap(km, t), t * TileSize, km / NumTiles, t / NumTiles)
                GetNormals CLng(k), CLng(t)
                Shade = .Normal1.Y
                .Vecs(1).Color = RGB(255 * Shade, 255 * Shade, 255 * Shade)
                .Vecs(2).Color = .Vecs(1).Color
                .Vecs(3).Color = .Vecs(1).Color
                Shade = .Normal2.Y
                .Vecs(4).Color = RGB(255 * Shade, 255 * Shade, 255 * Shade)
                .Vecs(5).Color = .Vecs(4).Color
                .Vecs(6).Color = .Vecs(4).Color
            End With
        Next t
    Next k
End Function

Public Function Render()
    D3DDevice.SetStreamSource 0, mVB, VertexSizeInBytes
    D3DDevice.SetVertexShader D3DFVF_VERTEX
    D3DDevice.SetTexture 0, Texture
    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, NumVertices / 3
End Function

Private Function GetNormals(DX As Long, DY As Long)
    Dim V1 As D3DVECTOR
    Dim V2 As D3DVECTOR
    With Tiles(DX, DY)
        V1 = VectorSubtract(.Vecs(3).Position, .Vecs(2).Position)
        V2 = VectorSubtract(.Vecs(3).Position, .Vecs(1).Position)
        .Normal1 = CrossProduct(V1, V2)
        If .Normal1.Y < 0 Then
            .Normal1 = MakeVector(-.Normal1.X, -.Normal1.Y, -.Normal1.Z)
        End If
        .Normal1 = Normalise(.Normal1)
        V1 = VectorSubtract(.Vecs(6).Position, .Vecs(1).Position)
        .Normal2 = CrossProduct(V1, V2)
        If .Normal2.Y < 0 Then
            .Normal2 = MakeVector(-.Normal2.X, -.Normal2.Y, -.Normal2.Z)
        End If
        .Normal2 = Normalise(.Normal2)
    End With
End Function

Private Function CollisionDetection(P As D3DVECTOR) As Boolean
    Dim CTileX As Single, CTileZ As Single
    Dim Temp As Single, TempP As D3DVECTOR
    CollisionDetection = False
    Temp = P.X / TileSize
    CTileX = Temp - Abs(Temp - CLng(Temp))
    Temp = P.Z / TileSize
    CTileZ = Temp - Abs(Temp - CLng(Temp))
    If CTileX <= NumTiles And CTileZ <= NumTiles And CTileX > 0 And CTileZ > 0 Then
        With Tiles(CTileX, CTileZ)
            TempP = MakeVector(P.X - .Vecs(1).Position.X, 0, P.Z - .Vecs(1).Position.Z)
            If TempP.X >= TempP.Z Then
                Temp = Abs(DotProduct(P, .Normal1) - DotProduct(.Vecs(1).Position, .Normal1))
                If Temp <= Acc Then
                    CollisionDetection = True
                End If
            Else
                Temp = Abs(DotProduct(P, .Normal2) - DotProduct(.Vecs(1).Position, .Normal2))
                If Temp <= Acc Then
                    CollisionDetection = True
                End If
            End If
        End With
    End If
End Function

Private Function CollisionDetection2(DPX As Integer, DPY As Integer, P As D3DVECTOR) As Boolean
    Dim Temp As Single, TempP As D3DVECTOR
    CollisionDetection2 = False
    With Tiles(DPX, DPY)
        TempP = MakeVector(P.X - .Vecs(1).Position.X, 0, P.Z - .Vecs(1).Position.Z)
        If TempP.X >= TempP.Z Then
            Temp = Abs(DotProduct(P, .Normal1) - DotProduct(.Vecs(1).Position, .Normal1))
            If Temp <= Acc Then
                If P.X >= .Vecs(1).Position.X And P.X <= .Vecs(2).Position.X Then
                    If P.Z >= .Vecs(1).Position.Z And P.Z <= .Vecs(3).Position.Z Then
                        If P.X - .Vecs(1).Position.X >= P.Z - .Vecs(1).Position.Z Then
                            CollisionDetection2 = True
                        End If
                    End If
                End If
            End If
        Else
            Temp = Abs(DotProduct(P, .Normal2) - DotProduct(.Vecs(1).Position, .Normal2))
            If Temp <= Acc Then
                If P.X >= .Vecs(1).Position.X And P.X <= .Vecs(2).Position.X Then
                    If P.Z >= .Vecs(1).Position.Z And P.Z <= .Vecs(3).Position.Z Then
                        If P.X - .Vecs(1).Position.X < P.Z - .Vecs(1).Position.Z Then
                            CollisionDetection2 = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Public Function RayCollision(PStart As D3DVECTOR, PEnd As D3DVECTOR, POut As D3DVECTOR) As Integer
    Dim Dir As D3DVECTOR
    Dim t As Single, LastT As Single, TempP As D3DVECTOR
    Dim CN As Single, Leng As Single, TVec As D3DVECTOR
    Dir = Normalise(VectorSubtract(PEnd, PStart))
    Leng = Dist(PEnd, PStart) - 0.5
    If Leng > 0 Then
        For kk = 1 To NumTiles
            For TT = 1 To NumTiles
                For n = 1 To 2
                    With Tiles(kk, TT)
                        If n = 1 Then CN = DotProduct(Dir, .Normal1)
                        If n = 2 Then CN = DotProduct(Dir, .Normal2)
                        If CN <> 0 Then
                            If n = 1 Then t = DotProduct(VectorSubtract(.Vecs(1).Position, PStart), .Normal1) / CN
                            If n = 2 Then t = DotProduct(VectorSubtract(.Vecs(1).Position, PStart), .Normal2) / CN
                            If t > 0 Then
                                If LastT > 0 Then
                                    If t < LastT And t < Leng Then
                                        TVec = MakeVector(Dir.X * t + PStart.X, Dir.Y * t + PStart.Y, Dir.Z * t + PStart.Z)
                                        If CollisionDetection2(CInt(kk), CInt(TT), TVec) = True Then
                                            LastT = t
                                            POut = TVec
                                            RayCollision = (kk * NumTiles) + TT
                                        End If
                                    End If
                                ElseIf t < Leng Then
                                    TVec = MakeVector(Dir.X * t + PStart.X, Dir.Y * t + PStart.Y, Dir.Z * t + PStart.Z)
                                    If CollisionDetection2(CInt(kk), CInt(TT), TVec) = True Then
                                        LastT = t
                                        POut = TVec
                                        RayCollision = (kk * NumTiles) + TT
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next n
            Next TT
        Next kk
    End If
End Function

Public Function RayCollisionHor(PStart As D3DVECTOR, PEnd As D3DVECTOR, POut As D3DVECTOR) As Boolean
    Dim Dir As D3DVECTOR, PTiles(1 To 2) As Tile, Coords(1 To 2, 1 To 2) As Integer
    Dim t As Single, LastT As Single, TempP As D3DVECTOR
    Dim CN As Single, Leng As Single, TVec As D3DVECTOR
    Dir = Normalise(VectorSubtract(PEnd, PStart))
    Leng = Dist(PEnd, PStart)
    Dim Temp As Single
    Temp = PStart.X / TileSize
    Coords(1, 1) = CLng(Temp)
    If Temp > Coords(1, 1) Then Coords(1, 1) = Coords(1, 1) + 1
    If Coords(1, 1) <= 0 Then
        While Coords(1, 1) <= 0
            Coords(1, 1) = Coords(1, 1) + NumTiles
        Wend
    ElseIf Coords(1, 1) > NumTiles Then
        While Coords(1, 1) > NumTiles
            Coords(1, 1) = Coords(1, 1) - NumTiles
        Wend
    End If
    Temp = PStart.Z / TileSize
    Coords(1, 2) = CLng(Temp)
    If Temp > Coords(1, 2) Then Coords(1, 2) = Coords(1, 2) + 1
    If Coords(1, 2) <= 0 Then
        While Coords(1, 2) <= 0
            Coords(1, 2) = Coords(1, 2) + NumTiles
        Wend
    ElseIf Coords(1, 2) > NumTiles Then
        While Coords(1, 2) > NumTiles
            Coords(1, 2) = Coords(1, 2) - NumTiles
        Wend
    End If
    Temp = PEnd.X / TileSize
    Coords(2, 1) = CLng(Temp)
    If Temp > Coords(2, 1) Then Coords(2, 1) = Coords(2, 1) + 1
    If Coords(2, 1) <= 0 Then
        While Coords(2, 1) <= 0
            Coords(2, 1) = Coords(2, 1) + NumTiles
        Wend
    ElseIf Coords(2, 1) > NumTiles Then
        While Coords(2, 1) > NumTiles
            Coords(2, 1) = Coords(2, 1) - NumTiles
        Wend
    End If
    Temp = PEnd.Z / TileSize
    Coords(2, 2) = CLng(Temp)
    If Temp > Coords(2, 2) Then Coords(2, 2) = Coords(2, 2) + 1
    If Coords(2, 2) <= 0 Then
        While Coords(2, 2) <= 0
            Coords(2, 2) = Coords(2, 2) + NumTiles
        Wend
    ElseIf Coords(2, 2) > NumTiles Then
        While Coords(2, 2) > NumTiles
            Coords(2, 2) = Coords(2, 2) - NumTiles
        Wend
    End If
    PTiles(1) = Tiles(Coords(1, 1), Coords(1, 2))
    PTiles(2) = Tiles(Coords(2, 1), Coords(2, 2))
    If Leng > 0 Then
        For kk = 1 To 2
            For n = 1 To 2
                With PTiles(kk)
                    If n = 1 Then CN = Dir.X * .Normal1.X + Dir.Z * .Normal1.Z
                    If n = 2 Then CN = Dir.X * .Normal2.X + Dir.Z * .Normal2.Z
                    If CN <> 0 Then
                        If n = 1 Then t = DotProduct(VectorSubtract(.Vecs(1).Position, PStart), .Normal1) / CN
                        If n = 2 Then t = DotProduct(VectorSubtract(.Vecs(1).Position, PStart), .Normal2) / CN
                        If t > 0 Then
                            If LastT > 0 Then
                                If t < LastT And t < Leng Then
                                    TVec = MakeVector(Dir.X * t + PStart.X, Dir.Y * t + PStart.Y, Dir.Z * t + PStart.Z)
                                    If CollisionDetection2(Coords(kk, 1), Coords(kk, 2), TVec) = True Then
                                        LastT = t
                                        POut = TVec
                                        RayCollisionHor = Coords(kk, 1) * NumTiles + Coords(kk, 2)
                                    End If
                                End If
                            ElseIf t < Leng Then
                                TVec = MakeVector(Dir.X * t + PStart.X, Dir.Y * t + PStart.Y, Dir.Z * t + PStart.Z)
                                If CollisionDetection2(Coords(kk, 1), Coords(kk, 2), TVec) = True Then
                                    LastT = t
                                    POut = TVec
                                    RayCollisionHor = Coords(kk, 1) * NumTiles + Coords(kk, 2)
                                End If
                            End If
                        End If
                    End If
                End With
            Next n
        Next kk
    End If
End Function

Public Function RayCollisionDown(PStart As D3DVECTOR, POut As D3DVECTOR) As Boolean
    Dim t As Single, CN As Single
    Dim Temp As Single, TempP As D3DVECTOR
    Dim SX As Long, SY As Long
    Dim StartP As D3DVECTOR
    POut = PStart
    StartP = PStart
    Temp = StartP.X / TileSize
    StandX = CLng(Temp)
    If Temp > StandX Then StandX = StandX + 1
    SX = StandX
    If SX <= 0 Then
        While SX <= 0
            SX = SX + NumTiles
            StartP.X = StartP.X + NumTiles * TileSize
        Wend
    ElseIf SX > NumTiles Then
        While SX > NumTiles
            SX = SX - NumTiles
            StartP.X = StartP.X - NumTiles * TileSize
        Wend
    End If
    Temp = StartP.Z / TileSize
    StandY = CLng(Temp)
    If Temp > StandY Then StandY = StandY + 1
    SY = StandY
    If SY <= 0 Then
        While SY <= 0
            SY = SY + NumTiles
            StartP.Z = StartP.Z + NumTiles * TileSize
        Wend
    ElseIf SY > NumTiles Then
        While SY > NumTiles
            SY = SY - NumTiles
            StartP.Z = StartP.Z - NumTiles * TileSize
        Wend
    End If
    With Tiles(SX, SY)
        TempP = MakeVector(StartP.X - .Vecs(1).Position.X, 0, StartP.Z - .Vecs(1).Position.Z)
        If TempP.X >= TempP.Z Then
            CN = -.Normal1.Y
            If CN <> 0 Then
                t = DotProduct(VectorSubtract(.Vecs(1).Position, StartP), .Normal1) / CN
                TempP = MakeVector(StartP.X, -t + StartP.Y, StartP.Z)
                Temp = Abs(DotProduct(TempP, .Normal1) - DotProduct(.Vecs(1).Position, .Normal1))
                If Temp <= Acc Then
                    POut.Y = TempP.Y
                    RayCollisionDown = True
                End If
            End If
        Else
            CN = -.Normal2.Y
            If CN <> 0 Then
                t = DotProduct(VectorSubtract(.Vecs(1).Position, StartP), .Normal2) / CN
                TempP = MakeVector(StartP.X, -t + StartP.Y, StartP.Z)
                Temp = Abs(DotProduct(TempP, .Normal2) - DotProduct(.Vecs(1).Position, .Normal2))
                If Temp <= Acc Then
                    POut.Y = TempP.Y
                    RayCollisionDown = True
                End If
            End If
        End If
    End With
End Function

Public Function Lighting(Shadows As Boolean)
    Dim CurrCol As RGBValue, HitP As D3DVECTOR, Res As Single
    If NumLights > 0 Then
        For k = 1 To NumTiles
            For t = 1 To NumTiles
                For ps = 1 To 6
                    CurrCol.R = 20
                    CurrCol.G = 20
                    CurrCol.B = 20
                    For l = 1 To NumLights
                        With Tiles(k, t)
                            If Lights(l).Type = LIGHT_POINT Then
                                Res = RayCollision(Lights(l).Pos, .Vecs(ps).Position, HitP)
                                If Res = 0 Then
                                    CurrCol.R = CurrCol.R + Lights(l).Color.R
                                    CurrCol.G = CurrCol.G + Lights(l).Color.G
                                    CurrCol.B = CurrCol.B + Lights(l).Color.B
                                End If
                            ElseIf Lights(l).Type = LIGHT_DIRECTION Then
                                HitP = MakeVector(.Vecs(ps).Position.X - Lights(l).Dir.X * 10000, _
                                                  .Vecs(ps).Position.Y - Lights(l).Dir.Y * 10000, _
                                                  .Vecs(ps).Position.Z - Lights(l).Dir.Z * 10000)
                                Res = RayCollision(HitP, .Vecs(ps).Position, HitP)
                                If Res = 0 Then
                                    Res = -DotProduct(Lights(l).Dir, .Normal1)
                                    If Res > 0 Then
                                        CurrCol.R = CurrCol.R + Lights(l).Color.R * Res
                                        CurrCol.G = CurrCol.G + Lights(l).Color.G * Res
                                        CurrCol.B = CurrCol.B + Lights(l).Color.B * Res
                                    End If
                                End If
                            Else
                                CurrCol.R = CurrCol.R + Lights(l).Color.R
                                CurrCol.G = CurrCol.G + Lights(l).Color.G
                                CurrCol.B = CurrCol.B + Lights(l).Color.B
                            End If
                        End With
                    Next l
                    NumLights = NumLights + 1
                    Tiles(k, t).Vecs(ps).Color = RGB(CurrCol.R / NumLights, CurrCol.G / NumLights, CurrCol.B / NumLights)
                    NumLights = NumLights - 1
                Next ps
            Next t
        Next k
    End If
End Function
