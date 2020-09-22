VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Tris() As Triangle
Private NumTris As Integer
Private Textures() As Direct3DTexture8
Private NumTex As Integer
Private Pos As D3DVECTOR

Private MinX As Single, MaxX As Single
Private MinZ As Single, MaxZ As Single
Private MinY As Single, MaxY As Single

Private Const AccB As Single = 20

Public Function Init(FPath As String)
    Dim TT As String, CT As Triangle, TTex As String, Col As Long, TP As D3DVECTOR
    Open App.Path + FPath For Input As #1
        While EOF(1) = False
            Input #1, TT
            If UCase(TT) = "T" Then
                NumTris = NumTris + 1
                ReDim Preserve Tris(1 To NumTris) As Triangle
                With Tris(NumTris)
                    Input #1, .Vecs(1).Position.X, .Vecs(1).Position.Y, .Vecs(1).Position.Z, _
                              .Vecs(2).Position.X, .Vecs(2).Position.Y, .Vecs(2).Position.Z, _
                              .Vecs(3).Position.X, .Vecs(3).Position.Y, .Vecs(3).Position.Z, _
                              .Vecs(1).tu, .Vecs(1).tv, .Vecs(2).tu, .Vecs(2).tv, .Vecs(3).tu, .Vecs(3).tv, _
                              .Tex
                End With
            ElseIf UCase(TT) = "M" Then
                Input #1, TTex
                NumTex = NumTex + 1
                ReDim Preserve Textures(1 To NumTex) As Direct3DTexture8
                Set Textures(NumTex) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + TTex)
            ElseIf UCase(TT) = "P" Then
                Input #1, Pos.X, Pos.Y, Pos.Z
            End If
        Wend
    Close #1
    
    Col = Scene.GetTCol(Pos)
    
    For k = 1 To NumTris
        Set Tris(k).mVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, 0, D3DFVF_VERTEX, D3DPOOL_DEFAULT)
        For t = 1 To 3
            With Tris(k).Vecs(t).Position
                .X = .X + Pos.X
                .Y = .Y + Pos.Y
                .Z = .Z + Pos.Z
            End With
            Tris(k).Vecs(t).Color = Col
        Next t
    Next k
    MinX = 1000000
    MinZ = 1000000
    MinY = 1000000
    MaxX = -1000000
    MaxY = -1000000
    MaxZ = -1000000
    
    For k = 1 To NumTris
        For t = 1 To 3
            MinX = Min(MinX, Tris(k).Vecs(t).Position.X)
            MinY = Min(MinY, Tris(k).Vecs(t).Position.Y)
            MinZ = Min(MinZ, Tris(k).Vecs(t).Position.Z)
            MaxX = Max(MaxX, Tris(k).Vecs(t).Position.X)
            MaxY = Max(MaxY, Tris(k).Vecs(t).Position.Y)
            MaxZ = Max(MaxZ, Tris(k).Vecs(t).Position.Z)
        Next t
    Next k
    
    ReloadVB
    
End Function

Public Function CheckCollision(PStart As D3DVECTOR, PEnd As D3DVECTOR, POut As D3DVECTOR) As Boolean
    POut = PEnd
    If PEnd.X + AccB >= MinX And PEnd.Y + AccB >= MinY And PEnd.Z + AccB >= MinZ Then
        If PEnd.X <= MaxX + AccB And PEnd.Y <= MaxY + AccB And PEnd.Z <= MaxZ + AccB Then
            POut = PStart
            If PStart.X <= MinX Then
                POut.X = MinX - AccB
            ElseIf PStart.X >= MaxX Then
                POut.X = MaxX + AccB
            Else
                POut.X = PEnd.X
            End If
            If PStart.Y <= MinY Then
                POut.Y = MinY - AccB
            ElseIf PStart.Y >= MaxY Then
                POut.Y = MaxY + AccB
            Else
                POut.Y = PEnd.Y
            End If
            If PStart.Z <= MinZ Then
                POut.Z = MinZ - AccB
            ElseIf PStart.Z >= MaxZ Then
                POut.Z = MaxZ + AccB
            Else
                POut.Z = PEnd.Z
            End If
            CheckCollision = True
        End If
    End If
End Function

Public Function Render()
    For k = 1 To NumTris
        D3DDevice.SetStreamSource 0, Tris(k).mVB, VertexSizeInBytes
        D3DDevice.SetVertexShader D3DFVF_VERTEX
        D3DDevice.SetTexture 0, Textures(Tris(k).Tex)
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 1
    Next k
End Function

Private Function ReloadVB()
    For k = 1 To NumTris
        D3DVertexBuffer8SetData Tris(k).mVB, 0, VertexSizeInBytes * 3, 0, Tris(k).Vecs(1)
    Next k
End Function
