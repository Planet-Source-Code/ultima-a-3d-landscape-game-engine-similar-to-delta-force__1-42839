Attribute VB_Name = "HelperModule_Declares"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Const Acc As Single = 0.1

Public Const Pi As Single = 3.14159265358979
Public Const D3DFVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Public Type VERTEX
    Position As D3DVECTOR
    Color As Long
    tu As Single
    tv As Single
End Type

Public Type Tile
    Vecs(1 To 6) As VERTEX
    Normal1 As D3DVECTOR
    Normal2 As D3DVECTOR
End Type

Public Type Triangle
    Vecs(1 To 3) As VERTEX
    Tex As Integer
    mVB As Direct3DVertexBuffer8
End Type

Public Type ATriangle
    Vecs(1 To 3) As VERTEX
    vVecs(1 To 3) As VERTEX
    Tex As Integer
    mVB As Direct3DVertexBuffer8
End Type

Public CARRY As Single
Public EyeDir As D3DVECTOR
Public EyePos As D3DVECTOR
Public VertexSizeInBytes As Single

