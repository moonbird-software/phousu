VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   Caption         =   "Paskahousu"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   225
   ClientWidth     =   10410
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   694
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   7
      Left            =   9000
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   0
      Top             =   6720
      Width           =   1065
      Visible         =   0   'False
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   17
      Left            =   7920
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   25
      Top             =   4920
      Width           =   1200
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   18
      Left            =   7920
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   24
      Top             =   3360
      Width           =   1200
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   19
      Left            =   7920
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   23
      Top             =   1800
      Width           =   1200
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1560
      Index           =   14
      Left            =   6360
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   22
      Top             =   1800
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1560
      Index           =   15
      Left            =   5160
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   21
      Top             =   1800
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1560
      Index           =   16
      Left            =   3960
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   20
      Top             =   1800
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   11
      Left            =   1680
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   19
      Top             =   4920
      Width           =   1200
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   12
      Left            =   1680
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   18
      Top             =   3360
      Width           =   1200
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   13
      Left            =   1680
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   17
      Top             =   1800
      Width           =   1200
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1560
      Index           =   10
      Left            =   6120
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   16
      Top             =   5280
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1560
      Index           =   9
      Left            =   4920
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   15
      Top             =   5280
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1560
      Index           =   8
      Left            =   3720
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   14
      Top             =   5280
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   0
      Left            =   5055
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   13
      Top             =   6975
      Width           =   585
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   1
      Left            =   120
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   12
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Index           =   3
      Left            =   9570
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   11
      Top             =   5400
      Width           =   735
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1545
      Index           =   4
      Left            =   2910
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   10
      Top             =   2602
      Width           =   1155
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   5
      Left            =   4320
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   9
      ToolTipText     =   "Sika"
      Top             =   2602
      Width           =   1605
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   2
      Left            =   6615
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   8
      Top             =   0
      Width           =   585
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1545
      Index           =   6
      Left            =   6000
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   6
      Top             =   2602
      Width           =   1155
      Visible         =   0   'False
   End
   Begin VB.Frame fraDebug 
      BackColor       =   &H00008000&
      Caption         =   "Säännöt"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
      Visible         =   0   'False
      Begin VB.CheckBox chkDebug 
         BackColor       =   &H00008000&
         Caption         =   "Juvelat"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkDebug 
         BackColor       =   &H00008000&
         Caption         =   "Vain 10"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkDebug 
         BackColor       =   &H00008000&
         Caption         =   "Kuvat"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   8010
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdAction 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3300
      TabIndex        =   7
      Top             =   4440
      Width           =   1935
      Visible         =   0   'False
   End
   Begin VB.Image imaTurn 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "Main.frx":5C12
      Top             =   2640
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imaTurn 
      Height          =   480
      Index           =   0
      Left            =   960
      Picture         =   "Main.frx":6854
      Top             =   2160
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Label lblPlayer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelaaja 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4020
      TabIndex        =   29
      Top             =   7440
      Width           =   885
   End
   Begin VB.Label lblPlayer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelaaja 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   240
      Width           =   780
   End
   Begin VB.Label lblPlayer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelaaja 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   7320
      TabIndex        =   27
      Top             =   0
      Width           =   780
   End
   Begin VB.Label lblPlayer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelaaja 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   9525
      TabIndex        =   26
      Top             =   6240
      Width           =   780
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Peli"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&Uusi peli"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameNetwork 
         Caption         =   "&Verkkopeli..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGame0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameSettings 
         Caption         =   "&Asetukset..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameScore 
         Caption         =   "&Pisteet..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuGameSound 
         Caption         =   "&Äänet"
      End
      Begin VB.Menu mnuGame1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameDemo 
         Caption         =   "&Demo"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGame2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "&Lopeta"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Ohje"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Ohjeen aiheet"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Tietoja..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDebug_Click(Index As Integer)
    If fraDebug.Visible Then
        Rules.HonoursOpen = -chkDebug(0).Value
        Rules.Only10s = -chkDebug(1).Value
        Rules.DeucesWild = -chkDebug(2).Value
    End If
End Sub

Private Sub cmdAction_Click()
    ActionClick
End Sub

Private Sub Form_Load()
    Me.Refresh
    GameInit
End Sub

Private Sub Form_Resize()
    FormMainResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GameUninit
End Sub

Private Sub mnuGameDemo_Click()
    GameDemo
End Sub
Private Sub mnuGameExit_Click()
    Form_Unload False
End Sub

Private Sub mnuGameNetwork_Click()
    GameNetwork
End Sub

Private Sub mnuGameNew_Click()
    GameNew
End Sub
Private Sub mnuGameScore_Click()
    GameScore True
End Sub

Private Sub mnuGameSettings_Click()
    GameSettings
End Sub

Private Sub mnuGameSound_Click()
    mnuGameSound.Checked = Not mnuGameSound.Checked
    Game.Sound = mnuGameSound.Checked
    SaveSettings
End Sub

Private Sub mnuHelpAbout_Click()
    HelpAbout
End Sub

Private Sub mnuHelpContents_Click()
    HelpContents
End Sub

Private Sub picDeck_KeyPress(Index As Integer, KeyAscii As Integer)
    DeckKeyPress Index, Deck(IDD_DEALER), Deck(IDD_TRICK), Deck(IDD_USER), KeyAscii
End Sub
Private Sub picDeck_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    
    DeckClick Deck(Index), Deck(IDD_DEALER), Deck(IDD_TRICK), Deck(IDD_USER), GetCardIndex(Deck(Index), X, Y)
End Sub

