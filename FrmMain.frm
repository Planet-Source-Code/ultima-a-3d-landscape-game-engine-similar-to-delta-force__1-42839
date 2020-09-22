VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16815
   BeginProperty Font 
      Name            =   "Verdana Ref"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   857
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1121
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fLoading 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   11040
      Visible         =   0   'False
      Width           =   16815
      Begin VB.Label lLoading 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   16815
      End
   End
   Begin VB.Frame fMenu 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   10935
      Left            =   7320
      TabIndex        =   9
      Top             =   0
      Width           =   9495
      Begin VB.CheckBox cShadows 
         BackColor       =   &H00000000&
         Caption         =   "Use Shadows"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   7800
         Width           =   7215
      End
      Begin VB.CommandButton cReturn 
         BackColor       =   &H00808080&
         Caption         =   "Start Game"
         Height          =   495
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7080
         Width           =   7815
      End
      Begin VB.CommandButton cExit 
         BackColor       =   &H00808080&
         Caption         =   "Exit"
         Height          =   495
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   8400
         Width           =   7815
      End
      Begin VB.Image BarKnop 
         Height          =   1260
         Left            =   0
         Picture         =   "FrmMain.frx":0000
         Top             =   9720
         Width           =   1260
      End
      Begin VB.Image BarBot 
         Height          =   735
         Left            =   600
         Picture         =   "FrmMain.frx":04E9
         Top             =   10200
         Width           =   9000
      End
      Begin VB.Image LeftBar 
         Height          =   10545
         Left            =   0
         Picture         =   "FrmMain.frx":1C2C
         Top             =   0
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   4215
         Left            =   1200
         Top             =   1440
         Width           =   7815
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         Height          =   615
         Left            =   1200
         Top             =   7680
         Width           =   7815
      End
      Begin VB.Label lMissionName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Operation"
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   960
         Width           =   7815
      End
      Begin VB.Label lMissionBrief 
         BackColor       =   &H00000000&
         Caption         =   "-----"
         ForeColor       =   &H00E0E0E0&
         Height          =   4215
         Left            =   1200
         TabIndex        =   11
         Top             =   1440
         Width           =   7815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.PictureBox Bar 
      BackColor       =   &H00000000&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   16755
      TabIndex        =   5
      Top             =   11640
      Visible         =   0   'False
      Width           =   16815
      Begin VB.Label lHealth 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lBullets 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   12600
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lClips 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   14760
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image iClip 
         Height          =   870
         Left            =   15960
         Picture         =   "FrmMain.frx":2F9A
         Top             =   120
         Width           =   600
      End
      Begin VB.Image iHealth 
         Height          =   870
         Left            =   1920
         Picture         =   "FrmMain.frx":34A4
         Top             =   120
         Width           =   870
      End
      Begin VB.Image iBullets 
         Height          =   870
         Left            =   14160
         Picture         =   "FrmMain.frx":3C66
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox GameView 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11580
      Left            =   0
      ScaleHeight     =   772
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1121
      TabIndex        =   4
      Top             =   0
      Width           =   16815
   End
   Begin VB.PictureBox HMap 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   16320
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      Top             =   12720
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastX As Single, LastY As Single
Dim InZ As Boolean

Dim Started As Boolean

Private Const DegToRad As Single = 3.14159275180032 / 180
Private Const MouseSens As Single = 10

Private Sub cExit_Click()
    MenuC = 2
    DoEvents
    KillApp
    End
End Sub

Private Sub cReturn_Click()
    MenuC = 1
    SetCursorPos LastX + (FrmMain.Left / Screen.TwipsPerPixelX), LastY + (FrmMain.Top / Screen.TwipsPerPixelY)
    If Started = False Then
        Bar.Visible = True
        fMenu.Visible = False
        Started = True
        cReturn.Caption = "Return to Game"
        ShowCursor False
        fLoading.Visible = True
        DoEvents
        GameView.SetFocus
        StartUpGame
    End If
End Sub

Private Sub Form_Load()
    DoEvents
    Me.Show
    
    LoadMBrief
    
    LastX = 300
    LastY = 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillApp
    End
End Sub

Private Sub GameView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Zoom = ZoomIn Then
            Zoom = ZoomOut
            ZoomS.PlaySound False
        Else
            Zoom = ZoomIn
            ZoomS.PlaySound False
        End If
    End If
    If Button = 1 Then
        Weapons(CurrWeapon).Shoot
    End If
End Sub

Private Sub GameView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MenuC = -1 Then
        RotateCamera (LastX - X) * -DegToRad / MouseSens, (LastY - Y) * -DegToRad / MouseSens
        If X <> LastX Or Y <> LastY Then
            SetCursorPos LastX + (FrmMain.Left / Screen.TwipsPerPixelX), LastY + (FrmMain.Top / Screen.TwipsPerPixelY)
        End If
    End If
End Sub

