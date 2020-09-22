Attribute VB_Name = "LightingEngine"
Public Const LIGHT_POINT As Single = 1
Public Const LIGHT_DIRECTION As Single = 2
Public Const LIGHT_AMBIENT As Single = 3

Public Const MaxRange As Single = 1000

Public Type RGBValue
    R As Single
    G As Single
    B As Single
End Type

Public Type Light
    Pos As D3DVECTOR
    Dir As D3DVECTOR
    Color As RGBValue
    Type As Integer
End Type
    
Public Lights() As Light
Public NumLights As Integer
