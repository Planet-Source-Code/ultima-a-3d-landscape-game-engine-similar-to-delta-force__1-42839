VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EWater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const D3DFVF_WVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Private Type WVERTEX
    Position As D3DVECTOR
    Color As Long
    tu1 As Single
    tv1 As Single
End Type

Private Verts(1 To 4) As WVERTEX
Private SVerts(1 To 4) As WVERTEX

Private MainTexture As Direct3DTexture8

Public Function Initialise(MinX As Single, MinZ As Single, MaxX As Single, MaxZ As Single, Hei As Single, MPath As String)
    Dim Col As Long
    Verts(1) = MakeWVertex(MinX, Hei, MinZ, 0, 1)
    Verts(2) = MakeWVertex(MaxX, Hei, MinZ, 1, 1)
    Verts(3) = MakeWVertex(MinX, Hei, MaxZ, 0, 0)
    Verts(4) = MakeWVertex(MaxX, Hei, MaxZ, 1, 0)
    For k = 1 To 4
        Col = GetCol(Scene.GetTCol(Verts(k).Position)).R / 1.2
        Verts(k).Color = RGB(Col, Col, Col)
        SVerts(k) = Verts(k)
    Next k
    Set MainTexture = D3DX.CreateTextureFromFile(D3DDevice, MPath)
End Function

Private Function Update()
    For k = 1 To 4
        Verts(k).tu1 = Verts(k).tu1 + 0.0005
        Verts(k).tv1 = Verts(k).tv1 + 0.0005
        SVerts(k).tu1 = SVerts(k).tu1 - 0.0005
        SVerts(k).tv1 = SVerts(k).tv1 - 0.0005
    Next k
    If Verts(3).tu1 > 1 Then
        For k = 1 To 4
            Verts(k).tu1 = Verts(k).tu1 - 1
            SVerts(k).tu1 = SVerts(k).tu1 + 1
        Next k
    End If
    If Verts(3).tv1 > 1 Then
        For k = 1 To 4
            Verts(k).tv1 = Verts(k).tv1 - 1
            SVerts(k).tv1 = SVerts(k).tv1 + 1
        Next k
    End If
    If EyePos.X >= Verts(1).Position.X And EyePos.Z >= Verts(1).Position.Z Then
        If EyePos.X <= Verts(4).Position.X And EyePos.Z <= Verts(4).Position.Z Then
            If EyePos.Y + PHei <= Verts(1).Position.Y Then
                If InWater = False Then
                    WaterSo.PlaySound False
                    WTime = Timer
                Else
                    If Timer - WTime >= BreathT Then
                        WTime = WTime + 1
                        GetHit 5
                    End If
                End If
                InWater = True
                SetFog 200, RGB(WatCol, 0, 0), 2000
                Exit Function
            End If
        End If
    End If
    InWater = False
    SetFog 1000, MFGCOl
End Function

Public Function Render()
    Update
    
    D3DDevice.SetVertexShader D3DFVF_WVERTEX

    D3DDevice.SetTexture 0, MainTexture
    
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Verts(1), Len(Verts(1))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SVerts(1), Len(SVerts(1))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Function

Private Function MakeWVertex(X As Single, Y As Single, Z As Single, U As Single, V As Single) As WVERTEX
    With MakeWVertex
        .Position.X = X
        .Position.Y = Y
        .Position.Z = Z
        .tu1 = U
        .tv1 = V
    End With
End Function
