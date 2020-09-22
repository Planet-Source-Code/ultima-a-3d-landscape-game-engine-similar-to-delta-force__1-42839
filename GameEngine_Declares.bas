Attribute VB_Name = "GameEngine_Declares"
Public CurrentMission As Integer
Public MissionName() As String
Public MissionBrief() As String
Public NumMission As Integer

Public Health As Single
Public LastFired As Single

Public Weapons() As GWeapon
Public NumWeapons As Integer
Public CurrWeapon As Integer

Public InWater As Boolean
Public WTime As Long
Public WaterSo As ESound

Public Zoom As Single
Public ZoomS As ESound

Public HitSo As ESound

Public Const ZoomIn As Single = Pi / 10
Public Const ZoomOut As Single = Pi / 3
Public Const MFGCOl As Long = &H884400
Public Const WatCol As Byte = 255
Public Const BreathT As Long = 10
