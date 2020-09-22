Attribute VB_Name = "GraphicsEngine_Declares"
Public Scene As ELandScape
Public SBox As ESkyBox
Public Objects() As EObject
Public NumObjects As Integer
Public Waters() As EWater
Public NumWaters As Integer
Public PHeight As Single
Public MoveF As Single
Public MoveS As Single
Public Finish As Boolean

Public MenuC As Integer

Public StandX As Long
Public StandY As Long
Public LastStandX As Long
Public LastStandY As Long
Public InAir As Single
Public OnFloor As Boolean
Public FallTime As Single

Public Const MoveSpeed As Single = 10
Public Const Gravity As Single = 5
Public Const PHei As Single = 80
Public Const CHei As Single = 30
Public Const JHei As Single = 160
Public Const JSpeed As Single = 10
