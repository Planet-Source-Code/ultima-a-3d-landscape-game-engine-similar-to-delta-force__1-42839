VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ESkyBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private LastX As Single, LastZ As Single, LastY As Single
Private Texture As Direct3DTexture8
Private STexture As Direct3DTexture8
Private Verts(1 To 30) As VERTEX
Private mVB As Direct3DVertexBuffer8

Private Const MX As Single = 20000
Private Const MY As Single = 5000
Private Const TY As Single = MY - 700
Private Const TX As Single = MX + 500
Private Const GM As Single = -10000

Public Function Init(FPath As String)
    Set Texture = D3DX.CreateTextureFromFile(D3DDevice, FPath)
    
    Verts(1) = MakeVertex(-TX, TY, -TX, 0.01, 0.01)
    Verts(2) = MakeVertex(TX, TY, -TX, 0.49, 0.01)
    Verts(3) = MakeVertex(TX, TY, TX, 0.49, 0.99)
    Verts(4) = Verts(1)
    Verts(5) = Verts(3)
    Verts(6) = MakeVertex(-TX, TY, TX, 0.01, 0.99)
    Verts(7) = MakeVertex(-MX, MY, -MX, 0.51, 0)
    Verts(8) = MakeVertex(-MX, GM, -MX, 0.51, 1)
    Verts(9) = MakeVertex(MX, GM, -MX, 0.9, 1)
    Verts(10) = Verts(9)
    Verts(11) = Verts(7)
    Verts(12) = MakeVertex(MX, MY, -MX, 0.9, 0)
    Verts(13) = Verts(12)
    Verts(14) = Verts(9)
    Verts(15) = MakeVertex(MX, GM, MX, 0.51, 1)
    Verts(16) = Verts(12)
    Verts(17) = Verts(15)
    Verts(18) = MakeVertex(MX, MY, MX, 0.51, 0)
    Verts(19) = Verts(18)
    Verts(20) = Verts(15)
    Verts(21) = MakeVertex(-MX, MY, MX, 0.9, 0)
    Verts(22) = Verts(15)
    Verts(23) = Verts(21)
    Verts(24) = MakeVertex(-MX, GM, MX, 0.9, 1)
    Verts(25) = Verts(24)
    Verts(26) = Verts(21)
    Verts(27) = Verts(7)
    Verts(28) = Verts(7)
    Verts(29) = Verts(24)
    Verts(30) = Verts(8)
    
    For k = 1 To 30
        Verts(k).Color = RGB(200, 200, 200)
    Next k
    
    Set mVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 30, 0, D3DFVF_VERTEX, D3DPOOL_DEFAULT)
    
    ReloadVB
End Function

Public Function Render()
    D3DDevice.SetStreamSource 0, mVB, VertexSizeInBytes
    D3DDevice.SetVertexShader D3DFVF_VERTEX
    D3DDevice.SetTexture 0, Texture
    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 10
End Function

Public Function Update()
    Dim DifX As Single, DifY As Single, DifZ As Single
    DifX = EyePos.X - LastX
    DifY = EyePos.Y - LastY
    DifZ = EyePos.Z - LastZ
    For k = 1 To 30
        With Verts(k).Position
            .X = .X + DifX
            .Y = .Y + DifY
            .Z = .Z + DifZ
        End With
    Next k
    LastX = EyePos.X
    LastY = EyePos.Y
    LastZ = EyePos.Z
    ReloadVB
End Function

Private Function ReloadVB()
    D3DVertexBuffer8SetData mVB, 0, VertexSizeInBytes * 30, 0, Verts(1)
End Function
