Attribute VB_Name = "EInputEngine"
Public DI As DirectInput8
Public DIDevice As DirectInputDevice8
Public DIState As DIKEYBOARDSTATE

Public Function InitDI()
    Set DI = DX.DirectInputCreate()
    Set DIDevice = DI.CreateDevice("GUID_SysKeyboard")
    
    DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIDevice.SetCooperativeLevel FrmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    
    DIDevice.Acquire
End Function

Public Function CheckKey()
    DIDevice.GetDeviceStateKeyboard DIState
    MoveF = 0
    MoveS = 0
    With DIState
        If .Key(DIK_P) <> 0 Then
            Debug.Print EyePos.X, EyePos.Z
        End If
        If .Key(DIK_W) <> 0 Then
            MoveF = MoveF + 1
        End If
        If .Key(DIK_S) <> 0 Then
            MoveF = MoveF - 1
        End If
        If .Key(DIK_A) <> 0 Then
            MoveS = MoveS - 1
        End If
        If .Key(DIK_D) <> 0 Then
            MoveS = MoveS + 1
        End If
        If .Key(DIK_LCONTROL) <> 0 Then
            If PHeight = PHei Then
                PHeight = CHei
            End If
        Else
            PHeight = PHei
        End If
        If .Key(DIK_SPACE) <> 0 Then
            If OnFloor = True Then
                InAir = JHei
            End If
        End If
        If .Key(DIK_ESCAPE) <> 0 Then
            If Finish = False Then
                Finish = True
            Else
                MenuC = 1
            End If
        End If
        If .Key(DIK_R) <> 0 Then
            Weapons(CurrWeapon).Reload
        End If
        For k = 1 To NumWeapons
            If .Key(k + 1) <> 0 Then
                If Weapons(k).UseThis = True Then
                    Weapons(CurrWeapon).StashAway
                    CurrWeapon = k
                    Weapons(k).UseThis
                    UpdateIBar
                End If
            End If
        Next k
    End With
End Function

