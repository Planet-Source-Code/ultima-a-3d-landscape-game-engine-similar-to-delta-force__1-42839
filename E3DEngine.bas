Attribute VB_Name = "E3DEngine"
Public DX As New DirectX8
Public D3DX As New D3DX8
Public D3D As Direct3D8
Public DS As DirectSound8
Public D3DDevice As Direct3DDevice8

Public Function Init(MisPath As String)

    InitHelp

    FrmMain.lLoading.Caption = "Loading...DirectX"
    DoEvents

    Set D3D = DX.Direct3DCreate()
    If D3D Is Nothing Then Exit Function
    
    Set DS = DX.DirectSoundCreate(vbNullString)
    DS.SetCooperativeLevel FrmMain.hWnd, DSSCL_PRIORITY
    
    Dim DispMode As D3DDISPLAYMODE
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    Dim D3DPP As D3DPRESENT_PARAMETERS
    
    With D3DPP
        .Windowed = 1
        .BackBufferHeight = 600
        .BackBufferWidth = 800
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        .BackBufferFormat = DispMode.Format
        .hDeviceWindow = FrmMain.GameView.hWnd
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
    End With
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, FrmMain.GameView.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
    If D3DDevice Is Nothing Then Exit Function
    
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
    D3DDevice.SetRenderState D3DRS_FOGENABLE, 1
    D3DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
    D3DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_LINEAR
    D3DDevice.SetRenderState D3DRS_RANGEFOGENABLE, 0
    
    SetFog 1000, MFGCOl
    
    InitGame
    
    FrmMain.lLoading.Caption = "Loading...Mission"
    DoEvents
    
    LoadMission MisPath
    
    EyePos = MakeVector(7 * 400, 440, 33 * 400)
    For k = 1 To 13
        Scene.UpdateLS -1, 1
    Next k
    EyeDir = MakeVector(0.6, 0, -1)
    EyeDir = Normalise(EyeDir)
    InitDI
    
    SetupMatrices

    
End Function

Public Function SetFog(Dist As Single, Col As Long, Optional MDis As Single = 10000)
    D3DDevice.SetRenderState D3DRS_FOGSTART, FToDW(Dist)
    D3DDevice.SetRenderState D3DRS_FOGEND, FToDW(MDis)
    D3DDevice.SetRenderState D3DRS_FOGCOLOR, Col
End Function

Public Function FToDW(flo As Single) As Long
    Dim Buf As D3DXBuffer
    Set Buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData Buf, 0, 4, 1, flo
    D3DX.BufferGetData Buf, 0, 4, 1, FToDW
End Function

Public Function SetupMatrices()

    Dim matView As D3DMATRIX
    D3DXMatrixLookAtLH matView, MakeVector(EyePos.X, EyePos.Y + PHeight, EyePos.Z), _
        MakeVector(EyeDir.X + EyePos.X, EyeDir.Y + EyePos.Y + PHeight, EyePos.Z + EyeDir.Z), MakeVector(0, 1, 0)
    
    D3DDevice.SetTransform D3DTS_VIEW, matView

    Dim matProj As D3DMATRIX
    
    D3DXMatrixPerspectiveFovLH matProj, Zoom, 1, 1, 100000
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Function

Public Function RenderAll()
    SetupMatrices

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    
    D3DDevice.BeginScene
        
        D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
        D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE
        D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
        D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        
        RenderGameItems
        
    D3DDevice.EndScene
    
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function

Public Function KillApp()
    Finish = True
    Set DX = Nothing
    Set D3DDevice = Nothing
    Set D3D = Nothing
End Function
