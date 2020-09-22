Attribute VB_Name = "GameEngine_Functions"
Public Function LoadMBrief()
    LoadMissionBrief
    
    CurrentMission = 1
    
    FrmMain.lMissionName.Caption = MissionName(CurrentMission)
    FrmMain.lMissionBrief.Caption = MissionBrief(CurrentMission)
End Function

Public Function StartUpGame()
    Dim RG As Boolean
     
    Health = 100
       
    FrmMain.lLoading.Caption = "Loading...Mission"
    DoEvents
    
    Init App.Path + "\missions\mission1.mis"
    InitWeapons
    
    Set ZoomS = New ESound
    ZoomS.Create App.Path + "\sounds\Zoom.wav"
    Zoom = ZoomOut

    Set WaterSo = New ESound
    WaterSo.Create App.Path + "\sounds\water.wav"
    
    Set HitSo = New ESound
    HitSo.Create App.Path + "\sounds\hit.wav"
    
    CurrWeapon = 1
    Weapons(CurrWeapon).MakeAvailable
    Weapons(CurrWeapon).UseThis
    
    Weapons(2).MakeAvailable
    
    FrmMain.fLoading.Visible = False
    DoEvents
    UpdateIBar
ReturnGame:
    MenuC = -1
    
    Finish = False
    MainLoop
    
    RG = Menu
    If RG = True Then GoTo ReturnGame
    
    KillApp
    End
End Function

Public Function GetHit(Damage As Single)
    Health = Health - Damage
    HitSo.PlaySound False
    UpdateIBar
    If Health <= 0 Then
        KillApp
    End If
End Function

Public Function UpdateIBar()
    FrmMain.lBullets.Caption = Weapons(CurrWeapon).Bullets
    FrmMain.lClips.Caption = Weapons(CurrWeapon).Clips
    FrmMain.lHealth = Health
End Function

Public Function LoadNextMission()
    CurrentMission = CurrentMission + 1
    LoadMission App.Path + "\missions\mission" + CStr(CurrentMission) + ".mis"
End Function

Public Function LoadMissionBrief()
    Dim TT As String, CM As Integer, F As String
    Open App.Path + "\missions\mission.mde" For Input As #3
        While EOF(3) = False
            Input #3, TT
            If UCase(TT) = "N" Then
                CM = CM + 1
                NumMission = CM
                ReDim Preserve MissionName(1 To NumMission) As String
                ReDim Preserve MissionBrief(1 To NumMission) As String
                Input #3, MissionName(CM)
            ElseIf UCase(TT) = "D" Then
                Input #3, F
                MissionBrief(CM) = MissionBrief(CM) + " " + F
            End If
        Wend
    Close #3
End Function

Public Function InitWeapons()
    Dim TT As String, FP As String, MP As String, B As Long, C As Long, MB As Long, MC As Long
    Open App.Path + "\weapons\weapon.des" For Input As #4
        While EOF(4) = False
            Input #4, TT
            If UCase(TT) = "W" Then
                Input #4, MP, FP, B, C, MB, MC
                NumWeapons = NumWeapons + 1
                ReDim Preserve Weapons(1 To NumWeapons) As GWeapon
                Set Weapons(NumWeapons) = New GWeapon
                Weapons(NumWeapons).Init MP, FP, MB, MC
                Weapons(NumWeapons).Bullets = B
                Weapons(NumWeapons).Clips = C
            End If
        Wend
    Close #4
End Function

Public Function Menu() As Boolean
    ShowCursor True
    MenuC = 0
    FrmMain.fMenu.Visible = True
    While MenuC = 0
        DoEvents
        If MenuC = 1 Then
            Menu = True
            FrmMain.fMenu.Visible = False
            ShowCursor False
            Exit Function
        ElseIf MenuC = 2 Then
            Menu = False
            Exit Function
        End If
    Wend
End Function
