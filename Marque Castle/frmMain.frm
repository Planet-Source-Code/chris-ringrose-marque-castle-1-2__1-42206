VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H000080FF&
   Caption         =   "Marque Castle 1.2"
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   -225
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   15
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Marque Castle"
      Filter          =   "*.gav"
   End
   Begin VB.Frame fraRegistry 
      BackColor       =   &H00808080&
      Caption         =   "Registry:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Left            =   9690
      TabIndex        =   21
      Top             =   7185
      Visible         =   0   'False
      Width           =   6225
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   33
         Top             =   3960
         Width           =   1155
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register Now!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   3960
         Width           =   1365
      End
      Begin VB.CheckBox chkRecieveEmail 
         BackColor       =   &H00808080&
         Caption         =   "Recieve notification of free Marque Castle upgrades via e-mail ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   30
         Top             =   3060
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.TextBox txtRegistryNumber 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         MaxLength       =   15
         TabIndex        =   29
         Top             =   2610
         Width           =   2565
      End
      Begin VB.TextBox txtLastName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         MaxLength       =   25
         TabIndex        =   28
         Top             =   1950
         Width           =   2565
      End
      Begin VB.TextBox txtFirstName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         MaxLength       =   25
         TabIndex        =   27
         Top             =   1290
         Width           =   2565
      End
      Begin VB.TextBox txtEmailAddress 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         MaxLength       =   30
         TabIndex        =   26
         Top             =   630
         Width           =   2565
      End
      Begin VB.Image imgHelp 
         Height          =   330
         Index           =   2
         Left            =   5415
         Picture         =   "frmMain.frx":030A
         ToolTipText     =   "Help on Registration"
         Top             =   4365
         Width           =   780
      End
      Begin VB.Shape shpPercent 
         BackColor       =   &H00400000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   165
         Left            =   3300
         Top             =   4110
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label lblRegistryInfo 
         BackColor       =   &H00808080&
         Caption         =   $"frmMain.frx":06CE
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   3270
         TabIndex        =   31
         Top             =   615
         Width           =   2655
      End
      Begin VB.Label lblRegistryNumber 
         BackColor       =   &H00808080&
         Caption         =   "Registry Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2340
         Width           =   2565
      End
      Begin VB.Label lblLastName 
         BackColor       =   &H00808080&
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   2565
      End
      Begin VB.Label lblFirstName 
         BackColor       =   &H00808080&
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1020
         Width           =   2565
      End
      Begin VB.Label lblEmailAddress 
         BackColor       =   &H00808080&
         Caption         =   "E-mail Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   2565
      End
      Begin VB.Shape shpTotal 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   3270
         Top             =   4080
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.Frame fraItems 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Items:"
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   7320
      TabIndex        =   11
      Top             =   1650
      Visible         =   0   'False
      Width           =   1875
      Begin VB.Timer tmrCement 
         Enabled         =   0   'False
         Interval        =   65
         Left            =   870
         Top             =   570
      End
      Begin VB.Timer tmrKey 
         Enabled         =   0   'False
         Interval        =   65
         Left            =   810
         Top             =   300
      End
      Begin VB.Line linTwo 
         BorderColor     =   &H000040C0&
         X1              =   165
         X2              =   870
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line linOne 
         BorderColor     =   &H000040C0&
         X1              =   165
         X2              =   870
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line linItemBoarder 
         X1              =   0
         X2              =   1860
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label lblItemInventory 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Item Inventory:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   225
         Left            =   120
         TabIndex        =   37
         Top             =   90
         Width           =   1440
      End
      Begin VB.Image imgHelp 
         Height          =   330
         Index           =   1
         Left            =   990
         Picture         =   "frmMain.frx":084A
         ToolTipText     =   "Help on Items"
         Top             =   915
         Width           =   780
      End
      Begin VB.Shape shpItemBoarder02 
         BackColor       =   &H00004080&
         BackStyle       =   1  'Opaque
         Height          =   1065
         Left            =   1800
         Top             =   360
         Width           =   75
      End
      Begin VB.Shape shpItemBoarder01 
         BackColor       =   &H00000000&
         Height          =   1395
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   1800
      End
      Begin VB.Label lblTimes3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   150
         Left            =   450
         TabIndex        =   35
         Top             =   690
         Width           =   135
      End
      Begin VB.Label lblCementBagsNum 
         BackColor       =   &H000080FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   34
         ToolTipText     =   "Total Bags of Cement"
         Top             =   600
         Width           =   285
      End
      Begin VB.Image imgCement 
         Height          =   195
         Left            =   180
         Stretch         =   -1  'True
         ToolTipText     =   "Turns water into Concrete"
         Top             =   660
         Width           =   195
      End
      Begin VB.Image imgBomb 
         Height          =   195
         Left            =   450
         Stretch         =   -1  'True
         ToolTipText     =   "Press SPACE BAR to use.  Has a 1 block radius."
         Top             =   930
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgClock 
         Height          =   195
         Left            =   720
         Stretch         =   -1  'True
         ToolTipText     =   "Resets the Clock to 150"
         Top             =   930
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgBoots 
         Height          =   195
         Left            =   180
         Stretch         =   -1  'True
         ToolTipText     =   "Allows you to walk on Spikes."
         Top             =   930
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgKey 
         Height          =   195
         Left            =   180
         Stretch         =   -1  'True
         ToolTipText     =   "Opens Doors and Locked Blocks"
         Top             =   390
         Width           =   195
      End
      Begin VB.Label lblKeysNum 
         BackColor       =   &H000080FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   13
         ToolTipText     =   "Total number of Keys"
         Top             =   330
         Width           =   285
      End
      Begin VB.Label lblTimes02 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   150
         Left            =   450
         TabIndex        =   12
         Top             =   420
         Width           =   135
      End
   End
   Begin VB.Frame fraTimer 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   7875
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Timer tmrWatchTime 
         Enabled         =   0   'False
         Interval        =   350
         Left            =   330
         Top             =   300
      End
      Begin VB.Timer tmrTimer 
         Enabled         =   0   'False
         Interval        =   900
         Left            =   900
         Top             =   270
      End
      Begin VB.Label lblTimer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "150"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   750
         TabIndex        =   17
         Top             =   75
         Width           =   465
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Time Left:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   165
         Left            =   75
         TabIndex        =   16
         Top             =   135
         Width           =   660
      End
   End
   Begin VB.Frame fraAI 
      BackColor       =   &H000080FF&
      Caption         =   "AI Timers:"
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   240
      TabIndex        =   14
      Top             =   5190
      Visible         =   0   'False
      Width           =   1095
      Begin VB.Timer tmrDeathAI 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   570
         Top             =   270
      End
      Begin VB.Timer tmrDroneAI 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   90
         Top             =   270
      End
   End
   Begin VB.Frame fraDefeat 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   9705
      TabIndex        =   2
      Top             =   3630
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Shape shpDefeatBorder 
         BorderColor     =   &H000040C0&
         BorderWidth     =   4
         Height          =   1035
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   -15
         Width           =   5205
      End
      Begin VB.Label lblDefeat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "To restart the level, press Ctrl+R, or goto Options, Restart Level, or click here!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   630
         Width           =   4965
      End
      Begin VB.Label lblDefeat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "D E F E A T !"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   150
         Width           =   4965
      End
   End
   Begin VB.Frame fraPaused 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   9705
      TabIndex        =   0
      Top             =   4650
      Visible         =   0   'False
      Width           =   5505
      Begin VB.Label lblPaused 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "To resume, press F12, or goto Options, Resume Game, or click here!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   39
         Top             =   510
         Width           =   5070
      End
      Begin VB.Label lblPaused 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "G A M E  P A U S E D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   150
         Width           =   5340
      End
      Begin VB.Shape shpPausedBorder 
         BorderColor     =   &H000040C0&
         BorderWidth     =   4
         Height          =   885
         Left            =   -15
         Shape           =   4  'Rounded Rectangle
         Top             =   -15
         Width           =   5550
      End
   End
   Begin VB.Frame fraLevelInfo 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   420
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Timer tmrDeath 
         Enabled         =   0   'False
         Interval        =   25
         Left            =   1140
         Top             =   435
      End
      Begin VB.Timer tmrScore 
         Enabled         =   0   'False
         Interval        =   465
         Left            =   1440
         Top             =   1110
      End
      Begin VB.Label lblSteps 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0000000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   1
         Left            =   660
         TabIndex        =   45
         Top             =   825
         Width           =   945
      End
      Begin VB.Label lblScore 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0000000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   615
         TabIndex        =   18
         Top             =   990
         Width           =   990
      End
      Begin VB.Label lblScoreTitle 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   " Score:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   36
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label lblSteps 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   " Steps:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   150
         TabIndex        =   44
         Top             =   825
         Width           =   1440
      End
      Begin VB.Line linTopCorner 
         BorderColor     =   &H000080FF&
         Index           =   1
         X1              =   1605
         X2              =   1605
         Y1              =   795
         Y2              =   825
      End
      Begin VB.Line linTopCorner 
         BorderColor     =   &H000080FF&
         Index           =   0
         X1              =   135
         X2              =   135
         Y1              =   795
         Y2              =   825
      End
      Begin VB.Shape shpScores 
         BorderColor     =   &H000040C0&
         Height          =   405
         Left            =   135
         Top             =   810
         Width           =   1485
      End
      Begin VB.Line linLevelBoarder 
         X1              =   0
         X2              =   1740
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Shape shpLevelBoarder02 
         BackColor       =   &H00004080&
         BackStyle       =   1  'Opaque
         Height          =   1065
         Left            =   1740
         Top             =   210
         Width           =   75
      End
      Begin VB.Shape shpLevelBoarder01 
         BackColor       =   &H00000000&
         Height          =   1395
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   1755
      End
      Begin VB.Label lblLives 
         BackColor       =   &H000080FF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   975
         TabIndex        =   9
         Top             =   450
         Width           =   285
      End
      Begin VB.Label lblTimes01 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   150
         Left            =   825
         TabIndex        =   8
         Top             =   540
         Width           =   135
      End
      Begin VB.Image imgGeorge 
         Height          =   285
         Left            =   435
         Stretch         =   -1  'True
         Top             =   480
         Width           =   285
      End
      Begin VB.Label lblLevelTitle 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Level Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   465
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   1755
      End
   End
   Begin VB.Frame fraLevel 
      BackColor       =   &H0080C0FF&
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
      ForeColor       =   &H00000000&
      Height          =   5190
      Left            =   2355
      TabIndex        =   5
      Top             =   1650
      Width           =   4890
      Begin VB.Frame fraHighScore 
         BackColor       =   &H000080FF&
         Caption         =   "Level Number"
         ForeColor       =   &H00FFFFFF&
         Height          =   1875
         Left            =   795
         TabIndex        =   51
         Top             =   525
         Visible         =   0   'False
         Width           =   3285
         Begin VB.CommandButton cmdAction 
            Caption         =   "Update"
            Height          =   255
            Left            =   2250
            Picture         =   "frmMain.frx":0C0E
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   1440
            Width           =   885
         End
         Begin VB.TextBox txtCurrentScore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   210
            MaxLength       =   25
            TabIndex        =   52
            Text            =   "Player Name"
            ToolTipText     =   "Enter your name"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Label lblCongratulations 
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Congratulations!  Enter your name into the Best Time archives!"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   90
            TabIndex        =   71
            Top             =   1350
            Width           =   2055
         End
         Begin VB.Label lblScoreNum 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            Height          =   165
            Index           =   4
            Left            =   120
            TabIndex        =   70
            Top             =   1080
            Width           =   75
         End
         Begin VB.Label lblScoreNum 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            Height          =   165
            Index           =   3
            Left            =   120
            TabIndex        =   69
            Top             =   930
            Width           =   75
         End
         Begin VB.Label lblScoreNum 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   165
            Index           =   2
            Left            =   120
            TabIndex        =   68
            Top             =   780
            Width           =   75
         End
         Begin VB.Label lblScoreNum 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   165
            Index           =   1
            Left            =   120
            TabIndex        =   67
            Top             =   630
            Width           =   75
         End
         Begin VB.Label lblScoreNum 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   165
            Index           =   0
            Left            =   120
            TabIndex        =   66
            Top             =   480
            Width           =   75
         End
         Begin VB.Label lblTopScore 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player's Score"
            Height          =   165
            Index           =   4
            Left            =   2070
            TabIndex        =   65
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label lblTopScore 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player's Score"
            Height          =   165
            Index           =   3
            Left            =   2070
            TabIndex        =   64
            Top             =   930
            Width           =   1035
         End
         Begin VB.Label lblTopScore 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player's Score"
            Height          =   165
            Index           =   2
            Left            =   2070
            TabIndex        =   63
            Top             =   780
            Width           =   1035
         End
         Begin VB.Label lblTopScore 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player's Score"
            Height          =   165
            Index           =   1
            Left            =   2070
            TabIndex        =   62
            Top             =   630
            Width           =   1035
         End
         Begin VB.Label lblTopScore 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player's Score"
            Height          =   165
            Index           =   0
            Left            =   2070
            TabIndex        =   61
            ToolTipText     =   "Best time"
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label lblTopName 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player Name"
            Height          =   165
            Index           =   4
            Left            =   210
            TabIndex        =   60
            Top             =   1080
            Width           =   1785
         End
         Begin VB.Label lblTopName 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player Name"
            Height          =   165
            Index           =   3
            Left            =   210
            TabIndex        =   59
            Top             =   930
            Width           =   1785
         End
         Begin VB.Label lblTopName 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player Name"
            Height          =   165
            Index           =   2
            Left            =   210
            TabIndex        =   58
            Top             =   780
            Width           =   1785
         End
         Begin VB.Label lblTopName 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player Name"
            Height          =   165
            Index           =   1
            Left            =   210
            TabIndex        =   57
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label lblTopName 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player Name"
            Height          =   165
            Index           =   0
            Left            =   210
            TabIndex        =   56
            ToolTipText     =   "Best time"
            Top             =   480
            Width           =   1785
         End
         Begin VB.Label lblHighScoreTime 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Time"
            Height          =   165
            Left            =   2040
            TabIndex        =   55
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lblHighScorePlayer 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Player"
            Height          =   165
            Left            =   180
            TabIndex        =   54
            Top             =   300
            Width           =   1845
         End
      End
      Begin VB.Frame fraPictures 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   1380
         TabIndex        =   48
         Top             =   1875
         Visible         =   0   'False
         Width           =   2145
         Begin VB.Image pic 
            Height          =   120
            Index           =   0
            Left            =   90
            Stretch         =   -1  'True
            ToolTipText     =   "Grass"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   1
            Left            =   240
            Stretch         =   -1  'True
            ToolTipText     =   "Cement"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   2
            Left            =   390
            Stretch         =   -1  'True
            ToolTipText     =   "Key on Grass"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   3
            Left            =   540
            Stretch         =   -1  'True
            ToolTipText     =   "Key on Cement"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   4
            Left            =   690
            Stretch         =   -1  'True
            ToolTipText     =   "Tile on Grass"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   5
            Left            =   840
            Stretch         =   -1  'True
            ToolTipText     =   "Tile on Cement"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   6
            Left            =   990
            Stretch         =   -1  'True
            ToolTipText     =   "Spikes on Grass"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   7
            Left            =   1140
            Stretch         =   -1  'True
            ToolTipText     =   "Spikes on Cement"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   8
            Left            =   1290
            Stretch         =   -1  'True
            ToolTipText     =   "Locked Block"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   9
            Left            =   1440
            Stretch         =   -1  'True
            ToolTipText     =   "Block"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   10
            Left            =   1590
            Stretch         =   -1  'True
            ToolTipText     =   "Bush"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   11
            Left            =   1740
            Stretch         =   -1  'True
            ToolTipText     =   "Brick Wall"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   12
            Left            =   1890
            Stretch         =   -1  'True
            ToolTipText     =   "Wood"
            Top             =   210
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   13
            Left            =   90
            Stretch         =   -1  'True
            ToolTipText     =   "Water"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   15
            Left            =   390
            Stretch         =   -1  'True
            ToolTipText     =   "Top Left of Building"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   16
            Left            =   540
            Stretch         =   -1  'True
            ToolTipText     =   "Top Right of Building"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   17
            Left            =   690
            Stretch         =   -1  'True
            ToolTipText     =   "Bottom Left of Building"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   18
            Left            =   840
            Stretch         =   -1  'True
            ToolTipText     =   "Bottom Right of Building"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   19
            Left            =   990
            Stretch         =   -1  'True
            ToolTipText     =   "Left Wall of Building"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   20
            Left            =   1140
            Stretch         =   -1  'True
            ToolTipText     =   "Right Wall of Building"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   21
            Left            =   1290
            Stretch         =   -1  'True
            ToolTipText     =   "Top Wall of Building"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   22
            Left            =   1440
            Stretch         =   -1  'True
            ToolTipText     =   "Bottom Wall of Building"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   23
            Left            =   1590
            Stretch         =   -1  'True
            ToolTipText     =   "Locked Door"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   24
            Left            =   1740
            Stretch         =   -1  'True
            ToolTipText     =   "Top Left of Fountain"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   25
            Left            =   1890
            Stretch         =   -1  'True
            ToolTipText     =   "Top Centre of Fountain"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   26
            Left            =   90
            Stretch         =   -1  'True
            ToolTipText     =   "Top Right of Fountain"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   27
            Left            =   240
            Stretch         =   -1  'True
            ToolTipText     =   "Middle Left of Fountain"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   28
            Left            =   390
            Stretch         =   -1  'True
            ToolTipText     =   "Middle Centre of Fountain"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   29
            Left            =   540
            Stretch         =   -1  'True
            ToolTipText     =   "Middle Right of Fountain"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   30
            Left            =   690
            Stretch         =   -1  'True
            ToolTipText     =   "Bottom Left of Fountain"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   31
            Left            =   840
            Stretch         =   -1  'True
            ToolTipText     =   "Bottom Centre of Fountain"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   32
            Left            =   990
            Stretch         =   -1  'True
            ToolTipText     =   "Bottom Right of Fountain"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   33
            Left            =   1140
            Stretch         =   -1  'True
            ToolTipText     =   "Toggle Block OFF"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   34
            Left            =   1290
            Stretch         =   -1  'True
            ToolTipText     =   "Toggle Block ON"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   35
            Left            =   1440
            Stretch         =   -1  'True
            ToolTipText     =   "Metalic Boots"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   36
            Left            =   1590
            Stretch         =   -1  'True
            ToolTipText     =   "Clock"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   37
            Left            =   1740
            Stretch         =   -1  'True
            ToolTipText     =   "Bomb"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   91
            Left            =   90
            Stretch         =   -1  'True
            ToolTipText     =   "George Facing UP"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   93
            Left            =   390
            Stretch         =   -1  'True
            ToolTipText     =   "George Facing DOWN"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   94
            Left            =   540
            Stretch         =   -1  'True
            ToolTipText     =   "George Facing LEFT"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   95
            Left            =   690
            Stretch         =   -1  'True
            ToolTipText     =   "Dead George"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   991
            Left            =   1590
            Stretch         =   -1  'True
            ToolTipText     =   "Drone Mouse"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   992
            Left            =   1740
            Stretch         =   -1  'True
            ToolTipText     =   "Death Mouse"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   993
            Left            =   1890
            Stretch         =   -1  'True
            ToolTipText     =   "Dead Drone Mouse"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   38
            Left            =   1890
            Stretch         =   -1  'True
            ToolTipText     =   "Bag of Cement"
            Top             =   510
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   92
            Left            =   240
            Stretch         =   -1  'True
            ToolTipText     =   "George Facing RIGHT"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   14
            Left            =   240
            Stretch         =   -1  'True
            ToolTipText     =   "Wall"
            Top             =   360
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   893
            Left            =   1140
            Stretch         =   -1  'True
            ToolTipText     =   "Norman Facing DOWN"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   892
            Left            =   990
            Stretch         =   -1  'True
            ToolTipText     =   "Norman Facing RIGHT"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   891
            Left            =   840
            Stretch         =   -1  'True
            ToolTipText     =   "Norman Facing UP"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   894
            Left            =   1290
            Stretch         =   -1  'True
            ToolTipText     =   "Norman Facing LEFT"
            Top             =   660
            Width           =   120
         End
         Begin VB.Image pic 
            Height          =   120
            Index           =   895
            Left            =   1440
            Stretch         =   -1  'True
            ToolTipText     =   "Dead Norman"
            Top             =   660
            Width           =   120
         End
         Begin VB.Shape shpPicturesBoarder 
            Height          =   825
            Left            =   30
            Top             =   60
            Width           =   2085
         End
      End
      Begin VB.Frame fraSkinDir 
         BackColor       =   &H000080FF&
         Caption         =   "Skin Directory:"
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   720
         TabIndex        =   72
         Top             =   60
         Visible         =   0   'False
         Width           =   3465
         Begin VB.CommandButton cmdSkinDirCancel 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   2520
            Picture         =   "frmMain.frx":30BC
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   780
            Width           =   705
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            Height          =   255
            Left            =   2520
            Picture         =   "frmMain.frx":556A
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   510
            Width           =   705
         End
         Begin VB.TextBox txtSkinDir 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   630
            TabIndex        =   73
            Text            =   "original"
            Top             =   510
            Width           =   1815
         End
         Begin VB.Label lblRestore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000040C0&
            Caption         =   "Restore original"
            ForeColor       =   &H00C0E0FF&
            Height          =   165
            Left            =   150
            TabIndex        =   79
            Top             =   840
            Width           =   945
         End
         Begin VB.Image imgHelp 
            Height          =   330
            Index           =   0
            Left            =   2640
            Picture         =   "frmMain.frx":7A18
            ToolTipText     =   "Help on Skins"
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label lblSkinDirInfo 
            BackColor       =   &H000080FF&
            Caption         =   "The ""Skin Directory"" is the folder at where Marque Castle Loads its pictures from."
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   390
            TabIndex        =   78
            Top             =   1095
            Width           =   2505
         End
         Begin VB.Label lblSkinDir 
            BackColor       =   &H000080FF&
            Caption         =   "\Skins\"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   77
            Top             =   540
            Width           =   495
         End
         Begin VB.Label lblSkinDirTitle 
            BackColor       =   &H000080FF&
            Caption         =   "Please select a directory:"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   150
            TabIndex        =   76
            Top             =   300
            Width           =   1695
         End
      End
      Begin VB.Timer tmrExplosionSmall 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   -285
         Top             =   75
      End
      Begin VB.Timer tmrNewGame 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   450
         Top             =   4995
      End
      Begin VB.Frame fraBattleArena 
         BackColor       =   &H000080FF&
         Caption         =   "Battle Arena:"
         ForeColor       =   &H00FFFFFF&
         Height          =   2865
         Left            =   750
         TabIndex        =   89
         Top             =   1035
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cboStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            ItemData        =   "frmMain.frx":7DDC
            Left            =   180
            List            =   "frmMain.frx":7DF2
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   2220
            Width           =   1425
         End
         Begin VB.OptionButton optMultiLevel 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "Gwen Cottage"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   75
            MaskColor       =   &H00C0E0FF&
            TabIndex        =   95
            Top             =   1830
            Width           =   1515
         End
         Begin VB.OptionButton optMultiLevel 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "Castle Factory"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   75
            MaskColor       =   &H00C0E0FF&
            TabIndex        =   94
            Top             =   1590
            Width           =   1425
         End
         Begin VB.OptionButton optMultiLevel 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "Krilin Forest"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   75
            MaskColor       =   &H00C0E0FF&
            TabIndex        =   93
            Top             =   1350
            Width           =   1380
         End
         Begin VB.OptionButton optMultiLevel 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "Lake Marque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   75
            MaskColor       =   &H00C0E0FF&
            TabIndex        =   92
            Top             =   1110
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.CommandButton cmdCancelMulti 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   2520
            Picture         =   "frmMain.frx":7E45
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   2520
            Width           =   765
         End
         Begin VB.CommandButton cmdStartMulti 
            Caption         =   "Start!"
            Default         =   -1  'True
            Height          =   255
            Left            =   1680
            Picture         =   "frmMain.frx":A2F3
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label lblPlr2Rules 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Controlled with: W-Up; S-Down; A-Left; D-Right; Q to use Bomb"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   435
            Left            =   1860
            TabIndex        =   102
            Top             =   1920
            Width           =   1425
         End
         Begin VB.Label lblPlr1Rules 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Controlled with the arrow keys; SPACE BAR to use Bomb"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   435
            Left            =   1860
            TabIndex        =   101
            Top             =   1200
            Width           =   1425
         End
         Begin VB.Label lblPlayerTwo 
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Player Two: Norman"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1830
            TabIndex        =   100
            Top             =   1740
            Width           =   1530
         End
         Begin VB.Label lblPlayerOne 
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Player One: George"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1830
            TabIndex        =   99
            Top             =   1020
            Width           =   1485
         End
         Begin VB.Image imgPlayer 
            Height          =   225
            Index           =   1
            Left            =   1590
            Stretch         =   -1  'True
            Top             =   1740
            Width           =   225
         End
         Begin VB.Image imgPlayer 
            Height          =   225
            Index           =   0
            Left            =   1590
            Stretch         =   -1  'True
            Top             =   1020
            Width           =   225
         End
         Begin VB.Label lblMultiLevelTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Select Level:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   60
            TabIndex        =   98
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label lblMultiIntro 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "George and Norman duke it out in a one-on-one death match in the Marque Castle Battle Arena!"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   150
            TabIndex        =   97
            Top             =   195
            Width           =   3030
         End
      End
      Begin VB.Frame fraCredits 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   390
         TabIndex        =   87
         Top             =   3945
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Timer tmrEnding 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   -135
            Top             =   -135
         End
         Begin VB.Label lblCredits 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   825
            Left            =   0
            TabIndex        =   88
            Top             =   0
            Width           =   4095
         End
      End
      Begin VB.CommandButton cmdBegin02 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Begin"
         Height          =   285
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   4515
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Frame fraLoading 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   810
         TabIndex        =   85
         Top             =   4620
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.Frame fraMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2040
         Left            =   660
         TabIndex        =   80
         Top             =   1395
         Visible         =   0   'False
         Width           =   3600
         Begin VB.CommandButton cmdBegin 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "Begin"
            Height          =   285
            Left            =   2580
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   1620
            Width           =   915
         End
         Begin VB.TextBox txtMessage 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1275
            Left            =   75
            Locked          =   -1  'True
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   81
            TabStop         =   0   'False
            Text            =   "frmMain.frx":C7A1
            Top             =   300
            Width           =   3465
         End
         Begin VB.Label lblCreator 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   75
            TabIndex        =   84
            Top             =   1650
            Visible         =   0   'False
            Width           =   2310
         End
         Begin VB.Shape shpBoarder 
            BorderColor     =   &H00000000&
            BorderStyle     =   3  'Dot
            Height          =   2040
            Left            =   0
            Top             =   0
            Width           =   3600
         End
         Begin VB.Label lblLevelNum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "Level Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   15
            TabIndex        =   83
            Top             =   0
            Width           =   3585
         End
      End
      Begin VB.Timer tmrHideItemInfo 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   4470
         Top             =   5025
      End
      Begin VB.Timer tmrPts 
         Enabled         =   0   'False
         Interval        =   1100
         Left            =   -315
         Top             =   4020
      End
      Begin VB.Timer tmrExplosion 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   -180
         Top             =   -180
      End
      Begin VB.Label lblSecret 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "When the loading bar is still visible on the splash screen, quickly type ""gavannon123"" (without the quotes)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Index           =   1
         Left            =   390
         TabIndex        =   103
         Top             =   3255
         Visible         =   0   'False
         Width           =   4155
      End
      Begin VB.Label lblSecret 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   2
         Left            =   4170
         TabIndex        =   49
         Top             =   2850
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Image imgExplosion 
         Height          =   720
         Left            =   45
         Stretch         =   -1  'True
         Top             =   180
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image imgExplosionSmall 
         Height          =   390
         Left            =   195
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Shape shpBoarder01 
         Height          =   4860
         Left            =   0
         Top             =   150
         Width           =   4890
      End
      Begin VB.Label lblSecret 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Secret:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   945
         Index           =   0
         Left            =   225
         TabIndex        =   50
         Top             =   2835
         Visible         =   0   'False
         Width           =   4470
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   59
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Label lblItemInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   47
         Top             =   4980
         Width           =   4815
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "+ Speed up Drone/Death Mice, - Slow down Drone/Death Mice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   75
         TabIndex        =   46
         Top             =   -30
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.Label lblPts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   30
         TabIndex        =   38
         Top             =   4875
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   388
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   399
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   398
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   397
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   396
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   395
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   394
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   393
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   392
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   391
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   390
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   389
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   387
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   386
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   385
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   384
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   383
         Left            =   765
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   382
         Left            =   525
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   381
         Left            =   285
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   380
         Left            =   45
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   379
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   378
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   377
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   376
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   375
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   374
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   373
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   372
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   371
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   370
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   369
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   368
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   367
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   366
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   365
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   364
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   363
         Left            =   765
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   362
         Left            =   525
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   361
         Left            =   285
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   360
         Left            =   45
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   359
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   358
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   357
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   356
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   355
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   354
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   353
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   352
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   351
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   350
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   349
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   348
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   347
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   346
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   345
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   344
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   343
         Left            =   765
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   342
         Left            =   525
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   341
         Left            =   285
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   340
         Left            =   45
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   339
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   338
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   337
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   336
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   335
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   334
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   333
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   332
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   331
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   330
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   329
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   328
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   327
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   326
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   325
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   324
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   323
         Left            =   765
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   322
         Left            =   525
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   321
         Left            =   285
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   320
         Left            =   45
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   319
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   318
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   317
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   316
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   315
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   314
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   313
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   312
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   311
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   310
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   309
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   308
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   307
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   306
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   305
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   304
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   303
         Left            =   765
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   302
         Left            =   525
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   301
         Left            =   285
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   300
         Left            =   45
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   299
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   298
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   297
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   296
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   295
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   294
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   293
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   292
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   291
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   290
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   289
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   288
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   287
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   286
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   285
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   284
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   283
         Left            =   765
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   282
         Left            =   525
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   281
         Left            =   285
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   280
         Left            =   45
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   279
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   278
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   277
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   276
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   275
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   274
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   273
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   272
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   271
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   270
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   269
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   268
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   267
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   266
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   265
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   264
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   263
         Left            =   765
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   262
         Left            =   525
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   261
         Left            =   285
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   260
         Left            =   45
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   259
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   258
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   257
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   256
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   255
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   254
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   253
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   252
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   251
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   250
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   249
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   248
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   247
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   246
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   245
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   244
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   243
         Left            =   765
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   242
         Left            =   525
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   241
         Left            =   285
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   240
         Left            =   45
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   239
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   238
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   237
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   236
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   235
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   234
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   233
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   232
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   231
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   230
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   229
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   228
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   227
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   226
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   225
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   224
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   223
         Left            =   765
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   222
         Left            =   525
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   221
         Left            =   285
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   220
         Left            =   45
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   219
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   218
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   217
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   216
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   215
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   214
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   213
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   212
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   211
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   210
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   209
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   208
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   207
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   206
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   205
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   204
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   203
         Left            =   765
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   202
         Left            =   525
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   201
         Left            =   285
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   200
         Left            =   45
         Stretch         =   -1  'True
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   199
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   198
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   197
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   196
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   195
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   194
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   193
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   192
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   191
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   190
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   189
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   188
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   187
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   186
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   185
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   184
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   183
         Left            =   765
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   182
         Left            =   525
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   181
         Left            =   285
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   180
         Left            =   45
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   179
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   178
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   177
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   176
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   175
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   174
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   173
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   172
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   171
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   170
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   169
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   168
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   167
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   166
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   165
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   164
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   163
         Left            =   765
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   162
         Left            =   525
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   161
         Left            =   285
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   160
         Left            =   45
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   159
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   158
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   157
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   156
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   155
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   154
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   153
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   152
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   151
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   150
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   149
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   148
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   147
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   146
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   145
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   144
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   143
         Left            =   765
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   142
         Left            =   525
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   141
         Left            =   285
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   140
         Left            =   45
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   139
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   138
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   137
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   136
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   135
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   134
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   133
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   132
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   131
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   130
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   129
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   128
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   127
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   126
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   125
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   124
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   123
         Left            =   765
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   122
         Left            =   525
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   121
         Left            =   285
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   120
         Left            =   45
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   119
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   118
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   117
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   116
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   115
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   114
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   113
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   112
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   111
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   110
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   109
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   108
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   107
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   106
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   105
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   104
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   103
         Left            =   765
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   102
         Left            =   525
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   101
         Left            =   285
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   100
         Left            =   45
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   99
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   98
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   97
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   96
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   95
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   94
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   93
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   92
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   91
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   90
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   89
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   88
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   87
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   86
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   85
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   84
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   83
         Left            =   765
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   82
         Left            =   525
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   81
         Left            =   285
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   80
         Left            =   45
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   79
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   78
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   77
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   76
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   75
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   74
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   73
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   72
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   71
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   70
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   69
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   68
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   67
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   66
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   65
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   64
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   63
         Left            =   765
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   62
         Left            =   525
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   61
         Left            =   285
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   60
         Left            =   45
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   58
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   57
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   56
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   55
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   54
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   53
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   52
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   51
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   50
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   49
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   48
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   47
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   46
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   45
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   44
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   43
         Left            =   765
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   42
         Left            =   525
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   41
         Left            =   285
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   40
         Left            =   45
         Stretch         =   -1  'True
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   39
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   38
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   37
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   36
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   35
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   34
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   33
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   32
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   31
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   30
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   29
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   28
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   27
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   26
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   25
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   24
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   23
         Left            =   765
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   22
         Left            =   525
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   21
         Left            =   285
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   20
         Left            =   45
         Stretch         =   -1  'True
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   19
         Left            =   4605
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   18
         Left            =   4365
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   17
         Left            =   4125
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   16
         Left            =   3885
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   15
         Left            =   3645
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   14
         Left            =   3405
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   13
         Left            =   3165
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   12
         Left            =   2925
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   11
         Left            =   2685
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   10
         Left            =   2445
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   9
         Left            =   2205
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   8
         Left            =   1965
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   7
         Left            =   1725
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   6
         Left            =   1485
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   5
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   4
         Left            =   1005
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   3
         Left            =   765
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   2
         Left            =   525
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   1
         Left            =   285
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgMap 
         Height          =   240
         Index           =   0
         Left            =   45
         Stretch         =   -1  'True
         Top             =   180
         Width           =   240
      End
   End
   Begin VB.OLE GavannonCom 
      Appearance      =   0  'Flat
      AutoActivate    =   0  'Manual
      BackColor       =   &H0080C0FF&
      Class           =   "Package"
      Height          =   780
      Left            =   30
      OleObjectBlob   =   "frmMain.frx":C7D3
      SourceDoc       =   "C:\My Programs\Marque Castle\Gavannon.com.url"
      TabIndex        =   43
      Top             =   15
      UpdateOptions   =   2  'Manual
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Click File, New Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   240
      Index           =   1
      Left            =   2070
      TabIndex        =   42
      Top             =   900
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "The End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   750
      Index           =   0
      Left            =   1875
      TabIndex        =   41
      Top             =   255
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Label lblCheater 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheater!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   225
      Left            =   30
      TabIndex        =   40
      ToolTipText     =   "CheatMode.Enabled = True"
      Top             =   105
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Image imgBoarder02 
      Height          =   435
      Left            =   4725
      Picture         =   "frmMain.frx":E3EB
      ToolTipText     =   "http://www.gavannon.com/"
      Top             =   0
      Width           =   5040
   End
   Begin VB.Image imgTitle01 
      Height          =   1125
      Left            =   0
      Picture         =   "frmMain.frx":E94C
      ToolTipText     =   "© 2003 Chris Ringrose"
      Top             =   0
      Width           =   7650
   End
   Begin VB.Image imgTitleBoarder 
      Height          =   1125
      Left            =   45
      Picture         =   "frmMain.frx":F87A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9720
   End
   Begin VB.Label lblRegistered 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Register Here"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   0
      TabIndex        =   19
      ToolTipText     =   "Click to Register"
      Top             =   7485
      Width           =   1395
   End
   Begin VB.Label lblFilePath 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Left            =   75
      TabIndex        =   10
      Top             =   75
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblStrProperties 
      BackColor       =   &H00000000&
      Height          =   45
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLoadFile 
         Caption         =   "&Load Game"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuQuitGame 
         Caption         =   "&Quit Game"
      End
      Begin VB.Menu mnuSeperator01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveGame 
         Caption         =   "&Save Game"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveGameAs 
         Caption         =   "Save Game &As...."
      End
      Begin VB.Menu mnuSeperator02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateCustomGame 
         Caption         =   "&Create Custom Game...."
      End
      Begin VB.Menu mnuLoadCustomGame 
         Caption         =   "Play &Custom Game"
      End
      Begin VB.Menu mnuSeperator03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBattleArena 
         Caption         =   "Marque Battle Arena!"
      End
      Begin VB.Menu mnuSeperator04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPauseGame 
         Caption         =   "&Pause Game"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuRestartLevel 
         Caption         =   "&Restart Level"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuGotoLevel 
         Caption         =   "&Goto Level...."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSeperator05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSkins 
         Caption         =   "&Skins Directory...."
      End
      Begin VB.Menu mnuSeperator06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMusic 
         Caption         =   "&Music"
         Checked         =   -1  'True
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSound 
         Caption         =   "S&ound"
         Checked         =   -1  'True
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuItemInfo 
         Caption         =   "&Item Info"
         Checked         =   -1  'True
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuSeperator07 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBestTimes 
         Caption         =   "View &Best Times"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuGeneralHelp 
         Caption         =   "Marque Castle &Help"
         Begin VB.Menu mnuHelpSelect 
            Caption         =   "&Skins"
            Index           =   0
         End
         Begin VB.Menu mnuHelpSelect 
            Caption         =   "&Item Inventory"
            Index           =   1
         End
         Begin VB.Menu mnuHelpSelect 
            Caption         =   "&Registration"
            Index           =   2
         End
      End
      Begin VB.Menu mnuScenarioCreationHelp 
         Caption         =   "&Scenario Creation Help"
         Begin VB.Menu mnuHelpScenario 
            Caption         =   "Level ""&Rules"""
            Index           =   0
         End
         Begin VB.Menu mnuHelpScenario 
            Caption         =   "&General Help"
            Index           =   1
         End
      End
      Begin VB.Menu mnuTroubleshooting 
         Caption         =   "&Troubleshooting"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSeperator08 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnimestudiosMenu 
         Caption         =   "Gavasoft on the &Web"
         Begin VB.Menu mnuFreeDownloads 
            Caption         =   "Free Downloads"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuProductNews 
            Caption         =   "Product &News"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSeperator09 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAnimestudios 
            Caption         =   "http://www.gavannon.com/"
         End
      End
      Begin VB.Menu mnuSeperator10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About...."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Marque Castle v1.2
'frmMain
'Copyright © 2003 Chris Ringrose

'  [Note] - Refer to the GeekGuide.txt enclosed
'    1. All the code you see is derived from my own sweat and sleepless nights.
'         To make a long story short, it's all Copyright by me, so no matter
'         how flattering, please contact me before using any of my code.  I'll almost
'         always say yes, but that way it will be legal.

'    2. This was created under a short deadline, and I had only a year of
'         programming experience (a year being one class prior).  Allot of the
'         code is done inefficiently, and simply a quick solution to something
'         that seemed to be going wrong when debugging it.  (Hmm ... what if I
'         add 1 to i ... nope ... I'll subtract 1).
'       By no means is this a place to get good, rock solid functions.  It's
'         to get ideas, and help understand how you'd go about making a more
'         complex game in Visual Basic.  (Check Marque Castle Redux for better
'         functions when it is complete).

'    Sound effects for "Defeat" and "Water" made by me, others from from "Resident
'      Evil."  Theme Music from unknown ... sorry whoever ... ending music from
'      "EarthBound."  All other songs came with my Yamaha Midi XGPlayer.

'    Questions, comments, complaints, money?
'      You can contact me by e-mail at:
'      marque@gavannon.com


Option Explicit


Dim strGameSave As String


Private Sub cboStyle_GotFocus()
  cmdStartMulti.Default = True
End Sub

'  Closes the Top Scores Viewer
Private Sub cmdAction_Click()
  Dim i As Integer

    On Error Resume Next
    'Hides the Frame
        fraHighScore.Visible = False
    'Hides the Top Score Name entry
        txtCurrentScore.Visible = False
    'Makes a beeping sound
        If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
     'You've made a Top Time
          If cmdAction.Caption = "Update" Then
               'Updates strBTName
                   strBTName = txtCurrentScore.Text
               'Changes the stats
                    lblTopName(intTopTimePlace).Caption = txtCurrentScore.Text
                    lblTopScore(intTopTimePlace).BackColor = &HFF&
               'Resets the Input for a Top Time Name entry
                    txtCurrentScore.Text = "Player Name"
               'Updates the File Data
                    'Sets the value of strFile
                        If Mid$(strLevelFile, Len(strLevelFile) - 2, 3) = "lvl" Then strFile = App.Path & "\Scenarios\Lvl" & strLevel & ".bt"
                        If Mid$(strLevelFile, Len(strLevelFile) - 2, 3) = "cus" Then strFile = strBestTimesDir & "\Lvl" & strLevel & ".bt"
                    
                    'Closes all files opened (if any)
                         Close
                         Open strFile For Output As #1
                    'Updates the data
                         For i = 0 To 4
                              Write #1, lblTopName(i).Caption, Int(lblTopScore(i).Caption)
                         Next
                         Close
               'Loads the Next Level
                    'Makes sure it isn't the Last Level
                         If strNextLevel <> "Last" Then
                              'The name of the file to load:
                                    strFile = App.Path & "\Scenarios\" & strNextLevel
                                    If Mid$(strLevelFile, Len(strLevelFile) - 2, 3) = "lvl" Then strFile = App.Path & "\Scenarios\" & strNextLevel
                                    If Mid$(strLevelFile, Len(strLevelFile) - 2, 3) = "cus" Then strFile = strBestTimesDir & "\" & strNextLevel
                              'Sets the Score
                                   dblScore = dblScore + CDbl(lblTimer.Caption)
                                   ScoreUpdate
                              'Resets the Lives to 3 again
                                   If frmSplash.lblCheat.Visible = False Then
                                        intLivesNum = 5
                                   Else
                                        intLivesNum = 10
                                   End If
                              'Loads the Level
                                   LoadLevel
                              'Sets the Menu appropriately
                                   
                                   mnuQuitGame.Enabled = True
                                   mnuSaveGame.Enabled = True
                                   mnuSaveGameAs.Enabled = True
                                   mnuPauseGame.Enabled = True
                                   mnuRestartLevel.Enabled = True
                                   mnuBestTimes.Enabled = True
                              ' A Game is in progress
                                   booPlaying = True
                              'The game is not paused
                                   fraPaused.Visible = False
                                   fraDefeat.Visible = False
                                'Enables the menus
                                     frmMain.mnuFile.Enabled = True
                                     frmMain.mnuOptions.Enabled = True
                                     frmMain.mnuHelp.Enabled = True
                    'The Last Level
                        Else
                            booPlaying = False
                            frmMain.lblEnd(0).Visible = True
                            frmMain.lblEnd(1).Visible = True
                            frmMain.fraCredits.Visible = True
                            intCreditNum = 0
                            frmMain.tmrEnding.Enabled = True
                            frmMain.mnuFile.Enabled = False
                            frmMain.mnuOptions.Enabled = False
                            frmMain.mnuHelp.Enabled = False
                            If frmMain.mnuMusic.Checked = True Then frmSplash.medMidi.URL = App.Path & "\Music(5).mid"
                         End If
        Else
            'Re-enable all the menus
                frmMain.mnuFile.Enabled = True
                frmMain.mnuOptions.Enabled = True
                frmMain.mnuHelp.Enabled = True
        End If
End Sub


Private Sub cmdApply_Click()
  Dim a As String
  Dim i As Integer
  Dim iFileNum As Integer

  'Is this what you want to do?
  If MsgBox("This will require Marque Castle to restart." & vbNewLine & "Are you sure you want to change the skin and restart Marque Castle?", vbCritical + vbYesNo, "Marque Castle - Change Skin?") = vbYes Then
    'Fixes up the Directory
    a = txtSkinDir.Text
    
    'Makes it all small caps
    a = LCase(a)
    
    'Starts with a \
    If Mid$(a, 1, 1) = "\" Then a = Right(a, Len(a) - 1)
    
    'Ends with a \
    If Mid$(a, Len(a), 1) = "\" Then a = Left(a, Len(a) - 1)
    
    'Applies the changes
    txtSkinDir.Text = a
  
    'Sets the new Skin Directory
    strSkinDir = App.Path & "\skins\" & txtSkinDir
    
    'Saves the new Skin Directory to CusSet.opt
    'Updates the strProperties(0)
    strProperties(0) = txtSkinDir.Text
    
    'Ecrypts the Information
    For i = 0 To 4
      lblStrProperties.Caption = i
      EncryptedCusSet "strProperties(" & lblStrProperties.Caption & ")"
    Next
  
    'Closes any Files currently Open (If any)
    Close
    
    'Opens the Properties File
    strFile = App.Path & "\CusSet.opt"
    iFileNum = FreeFile
    Open strFile For Output As #iFileNum
    Write #iFileNum, strProperties(0), strProperties(1), strProperties(2), strProperties(3), strProperties(4)
    
    'Restart Marque Castle
    Shell App.Path & "\Marque Castle.exe", vbNormalFocus
    End
  End If
End Sub

'  Closes the Message to User on Startup Form
Private Sub cmdBegin_Click()
  Dim z As String

     If mnuSound.Checked = True Then
          'Makes a beeping sound
               PlaySound 0, App.Path & "\Beep.wav"
     End If

     'Hides the Message Form
          fraMessage.Visible = False

    'resets the steps
        dblSteps = 0
        frmMain.lblSteps(1).Caption = "000000000"

     'Changes booPlaying to True (You're now playing)
          booPlaying = True

     'Changes the Menu
          mnuFile.Enabled = True
          mnuOptions.Enabled = True
          mnuHelp.Enabled = True
          mnuQuitGame.Enabled = True
          mnuSaveGame.Enabled = True
          mnuSaveGameAs.Enabled = True
          mnuQuitGame.Enabled = True
          mnuPauseGame.Enabled = True
          mnuRestartLevel.Enabled = True
          mnuBestTimes.Enabled = True

     'Starts the timer
          lblTimer.Caption = "150"
          tmrTimer.Enabled = True

     'If there is a Drone Mouse in the Level
          If intDronePos >= 0 Then

               'There is a live Drone Mouse on the map
                    booDroneMouse = True
                    intDroneDir = 0
                    tmrDroneAI.Enabled = True

          End If
          
     'If there is a Death Mouse in the Level
          If intDeathPos >= 0 Then

               'There is a live Death Mouse on the map
                    blnDeathMouse = True
                    tmrDeathAI.Enabled = True

          End If

     'Sets the Form's Title appropriately
          frmMain.Caption = "Marque Castle - " & strLevelTitle

     'Starts the Mice's to move (If there are any)

          'If there is a Drone Mouse
               If booDroneMouse = True Then

                    'Sets the Drone Mouse Ground
                         intDroneGround = strDefaultGround

                    'Sets what way the Drone is facing (Right)
                         intDroneDir = 0

                    'Resets the number of times the Drone could not move
                         intDroneUnmove = 0

                    'Starts the Drone AI Timer
                         tmrDroneAI.Enabled = True

          'There isn't a Drone Mouse
               Else

                    'Disables the Drone AI Timer
                         tmrDroneAI.Enabled = False

               End If

    fraHighScore.Visible = False

End Sub


Private Sub cmdBegin02_Click()

     If mnuSound.Checked = True Then
          'Makes a beeping sound
               PlaySound 0, App.Path & "\Beep.wav"
     End If

     'Hides the Message Form
          fraMessage.Visible = False

    'resets the steps
        dblSteps = 0
        frmMain.lblSteps(1).Caption = "000000000"

     'Changes booPlaying to True (You're now playing)
          booPlaying = True

     'Changes the Menu
          mnuQuitGame.Enabled = True
          mnuSaveGame.Enabled = True
          mnuSaveGameAs.Enabled = True
          mnuQuitGame.Enabled = True
          mnuPauseGame.Enabled = True
          mnuRestartLevel.Enabled = True
          mnuBestTimes.Enabled = True

     'Shows the appropriate stats
          fraLevelInfo.Visible = True
          fraItems.Visible = True
          fraTimer.Visible = True

     'Checks if Music is Checked off
          If mnuMusic.Checked = True Then

               'Randomly Plays a MIDI Song
                    frmSplash.medMidi.URL = App.Path & "\Music(" & CStr(RndBetween(0, 4)) & ").mid"
'                    frmMain.Show

          End If

     'Starts the timer
          lblTimer.Caption = "150"
          tmrTimer.Enabled = True

     'If there is a Drone Mouse in the Level
          If intDronePos >= 0 Then

               'There is a live Drone Mouse on the map
                    booDroneMouse = True
                    intDroneDir = 0
                    tmrDroneAI.Enabled = True

          End If

     'If there is a Death Mouse in the Level
          If intDeathPos >= 0 Then

               'There is a live Death Mouse on the map
                    blnDeathMouse = True
                    tmrDeathAI.Enabled = True

          End If

     'Sets the Form's Title appropriately
          frmMain.Caption = "Marque Castle - " & strLevelTitle

     'Starts the Mice's to move (If there are any)

          'If there is a Drone Mouse
               If booDroneMouse = True Then

                    'Sets the Drone Mouse Ground
                         intDroneGround = strDefaultGround

                    'Sets what way the Drone is facing (Right)
                         intDroneDir = 0

                    'Resets the number of times the Drone could not move
                         intDroneUnmove = 0

                    'Starts the Drone AI Timer
                         tmrDroneAI.Enabled = True

          'There isn't a Drone Mouse
               Else

                    'Disables the Drone AI Timer
                         tmrDroneAI.Enabled = False

               End If

          cmdBegin02.Visible = False
         fraHighScore.Visible = False

End Sub



Private Sub cmdCancelMulti_Click()

     'Enables the Menu Commands
          mnuFile.Enabled = True
          mnuOptions.Enabled = True
          mnuHelp.Enabled = True

     'Hides the Battle Arena Frame
          fraBattleArena.Visible = False
    
     'Resets its options
          optMultiLevel(0).Value = True
          cboStyle.Text = "3 Point Win"

End Sub
'
''  Registers Marque Castle
'Private Sub cmdRegister_Click()
'  Dim i As Integer
'  Dim a As Integer
'  Dim b As Boolean
'
'     If mnuSound.Checked = True Then
'          'Makes a beeping sound
'               PlaySound 0, App.Path & "\Beep.wav"
'     End If
'
'     'Makes sure the E-mail Address contains an @ sign
'
'          'Counts the number of @'s
'               For i = 1 To Len(txtEmailAddress.Text)
'
'                    'Counts the @'s
'                         If Mid(txtEmailAddress.Text, i, 1) = "@" Then a = a + 1
'
'                    'Counts the spaces
'                         If Mid(txtEmailAddress.Text, i, 1) = " " Then a = a + 2
'
'                         If Mid(txtEmailAddress.Text, i, 1) = "." Then b = True
'
'               Next
'
'          'If more or less then one @
'               If a <> 1 And b = False And Len(txtEmailAddress.Text) < 5 Then
'
'                    'Displays error message
'                         MsgBox "Invalid E-mail Address.", vbCritical, "Marque Castle"
'
'                    'Exits the Sub
'                         Exit Sub
'
'               End If
'
'     'Makes sure that you input your name
'          If Len(txtFirstName) < 1 Or Len(txtLastName) < 1 Then
'
'               'Displays error message
'                    MsgBox "You must enter your full name to continue.", vbCritical, "Marque Castle"
'
'               'Exits the Sub
'                    Exit Sub
'
'          End If
'
'     'Makes sure that the Registry number is valid
'          If txtRegistryNumber <> "1586958755" And txtRegistryNumber <> "4562248679" _
'               And txtRegistryNumber <> "1384468790" And txtRegistryNumber <> "7801879321" _
'               And txtRegistryNumber <> "5493271209" And txtRegistryNumber <> "7254189307" _
'               And txtRegistryNumber <> "9462874215" And txtRegistryNumber <> "4781239855" _
'               And txtRegistryNumber <> "1597300489" And txtRegistryNumber <> "2687006811" _
'               And txtRegistryNumber <> "3687195482" And txtRegistryNumber <> "4589673009" Then
'
'                    'Displays error message
'                         MsgBox "Invalid Registration number.", vbCritical, "Marque Castle"
'
'                    'Exits the Sub
'                         Exit Sub
'
'          End If
'
'     'Shows the 'loading' total bar
'          shpTotal.Visible = True
'
'     'Resizes the Percent bar to zero
'          shpPercent.Width = 0
'
'     'Shows the percent bar
'          shpPercent.Visible = True
'
'     'Creates the 'loading' animation effect
'          For i = 0 To 2565 Step 5
'
'               'increases the size of the "percent bar"
'                    shpPercent.Width = i
'                    DoEvents
'
'          Next
'
'     'Save the information to CusSet.opt
'          'Updates the strProperties(3)
'               strProperties(3) = "registered"
'
'          'Ecripts the Information
'               For i = 0 To 4
'                    lblStrProperties.Caption = i
'                    EncryptedCusSet "strProperties(" & lblStrProperties.Caption & ")"
'               Next
'
'          'Closes any Files currently Open (If any)
'               Close
'
'          'Opens the Properties File
'               Open App.Path & "\CusSet.opt" For Output As #1
'               Write #1, strProperties(0), strProperties(1), strProperties(2), strProperties(3), _
'                    strProperties(4)
'
'          'Closes any Files currently Open
'               Close #1
'
'          'Creates the SEND_ME file
'               strFile = App.Path & "\SEND_ME.rgi"
'               Open strFile For Output As #1
'               Write #1, txtEmailAddress.Text, txtFirstName.Text, _
'                    txtLastName.Text, txtRegistryNumber.Text, chkRecieveEmail.Value
'
'     'Presents message boxes to the user, stating that registration is complete, as well as direction to further registration
'          MsgBox "Your Registraion of Marque Castle was complete!" & vbNewLine & vbNewLine & "You may now create and play your own Levels using the ''Scenario Creation Artist.''" & vbNewLine & "See Help for details.", vbInformation, "Marque Castle"
'          MsgBox "The Registration Information must be sent to Chris to activate your account." & vbNewLine & "Send the following File as an Attachment via E-mail to marque@gavannon.com:" & vbNewLine & vbNewLine & App.Path & "\SEND_ME.rgi''.", vbInformation, "Marque Castle"
'
'     'Closes the Registry Window
'          'Hides the Registry Window
'               fraRegistry.Visible = False
'               strProperties(3) = "registered"
'
'          'Changes the Border Style back to the original
'               lblRegistered.Visible = False
'
'          'Resets the Values in the Registry Window
'               txtEmailAddress.Text = ""
'               txtFirstName.Text = ""
'               txtLastName.Text = ""
'               txtRegistryNumber.Text = ""
'               chkRecieveEmail.Value = 1
'               shpTotal.Visible = False
'               shpPercent.Visible = False
'
'End Sub

Private Sub cmdSkinDirCancel_Click()

     'Hides the Skin Directory Frame
          fraSkinDir.Visible = False

     'Hides the Picture Display
          fraPictures.Visible = False

     'Decrypts strProperties(0)
          DecryptedCusSet "strProperties(0)"

     'Resets the value of the Skin Directory Text Box
          txtSkinDir.Text = strProperties(0)

     'Encrypts strProperties(0)
          EncryptedCusSet "strProperties(0)"

     'Enables all of the Menus
          mnuFile.Enabled = True
          mnuOptions.Enabled = True
          mnuHelp.Enabled = True

End Sub


Private Sub cmdStartMulti_Click()
  If fraPaused.Visible = True Then
    If MsgBox("This will Quit your current Game!", vbCritical + vbOKCancel, "Marque Castle") = vbCancel Then
      Exit Sub
    End If
  End If
  
  If cboStyle.Text = "" Then
    MsgBox "You must select the Game style first.", vbCritical, "Marque Castle"
    Exit Sub
  End If

  'Shows the Battle Arena
  frmBattleArena.Show 1
End Sub

'  Code on Startup
Private Sub Form_Load()


     'Adjusts the Title Boarder
          imgTitleBoarder.Width = frmMain.Width
     
     'Adjusts the Logo
          imgBoarder02.Left = frmMain.Width - imgBoarder02.Width

     'Centers each Object to fit the Form size

          'The Level display (fraLevel)
               fraLevel.Left = (frmMain.Width - fraLevel.Width) / 2
               fraLevel.Top = ((frmMain.Height - fraLevel.Height) / 2) + 500

          'The Level info Form (fraLevelInfo)
               fraLevelInfo.Left = (fraLevel.Left - fraLevelInfo.Width) / 2
               fraLevelInfo.Top = fraLevel.Top

          'The Items Menu (fraItems)
               fraItems.Left = (frmMain.Width + fraLevel.Left + fraLevel.Width) / 2 - fraItems.Width / 2
               fraItems.Top = fraLevel.Top

          'The Timer display (fraTimer)
               fraTimer.Left = (fraItems.Left + fraItems.Width) - fraTimer.Width
               fraTimer.Top = fraItems.Top + fraItems.Height + 35

          'The Registry Window
'               fraRegistry.Left = (frmMain.Width - fraRegistry.Width) / 2
'               fraRegistry.Top = (frmMain.Height - fraRegistry.Height) / 2

     'Sets the defaults for the Menu
          mnuQuitGame.Enabled = False
          mnuSaveGame.Enabled = False
          mnuSaveGameAs.Enabled = False
          mnuPauseGame.Enabled = False
          mnuRestartLevel.Enabled = False
          mnuBestTimes.Enabled = False

     'Sets the defaults for the Variables
          intKeysNum = 0
          booPlaying = False
          If frmSplash.lblCheat.Visible = False Then
               intLivesNum = 5
          Else
               intLivesNum = 10
          End If
          strBTName = "Player Name"
          cboStyle.Text = "3 Point Win"

     'Hides the "Paused" Frame
          fraPaused.Visible = False
          fraDefeat.Visible = False

End Sub


'  Upon Resizing the Form....
Private Sub Form_Resize()

    If Me.WindowState = 1 Then Exit Sub
    On Error Resume Next
     'Ensures that the New Form Size is appropriate
          'Too short
               If frmMain.Height < 7815 Then

                    'If it's too small, it resizes it
                         If Screen.Height <> 7200 Then frmMain.Height = 7815

               End If

          'Too thin
               If frmMain.Width < 9545 Then

                    'If it's too small, it resizes it
                         frmMain.Width = 9545

               End If


     'Adjusts the Title Boarder
          imgTitleBoarder.Width = frmMain.Width
     
     'Adjusts the Logo
          imgBoarder02.Left = frmMain.Width - imgBoarder02.Width
     'Centers each Object to fit the New Form size
          'The Level display (fraLevel)
               fraLevel.Left = (frmMain.ScaleWidth - fraLevel.Width) / 2
               fraLevel.Top = ((frmMain.ScaleHeight - fraLevel.Height) / 2) + 500
          'The Level info form (fraLevelInfo)
               fraLevelInfo.Left = (fraLevel.Left - fraLevelInfo.Width) / 2
               fraLevelInfo.Top = fraLevel.Top
          'The Items Menu (fraItems)
               fraItems.Left = (frmMain.Width + fraLevel.Left + fraLevel.Width) / 2 - fraItems.Width / 2
               fraItems.Top = fraLevel.Top
          'The Timer display (fraTimer)
               fraTimer.Left = (fraItems.Left + fraItems.Width) - fraTimer.Width
               fraTimer.Top = fraItems.Top + fraItems.Height + 35
          'The Defeat Title
               fraDefeat.Left = ((frmMain.Width - fraDefeat.Width) / 2) - 55
               fraDefeat.Top = frmMain.Height / 2
          'The Pause Title
               fraPaused.Left = ((frmMain.Width - fraPaused.Width) / 2) - 55
               fraPaused.Top = fraDefeat.Top
'          If lblRegistered.Visible = True Then
'               'The Registry Status
'                    lblRegistered.Left = 30
'                    lblRegistered.Top = lblMain.Height - lblRegistered.Height - 120
'          End If
          'The Registry Window
'               fraRegistry.Left = (frmMain.Width - fraRegistry.Width) / 2
'               fraRegistry.Top = (frmMain.Height - fraRegistry.Height) / 2
        If fraMessage.Visible = True Then
            fraMessage.Visible = False
            fraMessage.Visible = True
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Responce As VbMsgBoxResult
    
    'If playing a game, warn about quitting
    If booPlaying = True Then
        'Pause the game first
        If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
        fraPaused.Visible = True
        tmrDroneAI.Enabled = False
        tmrDeathAI.Enabled = False
        mnuPauseGame.Caption = "R&esume Game"
        fraPaused.Visible = True
        tmrTimer.Enabled = False

        'Ask the user if he/she would like to save first
        Responce = MsgBox("Would you like to Save your progress before exiting?", vbYesNoCancel, "Marque Castle")
        'Yes
            If Responce = vbYes Then
              SaveGame
              End
        'No
            ElseIf Responce = vbNo Then
              End
        'Cancel
            ElseIf Responce = vbCancel Then
              Cancel = 2
            End If
    Else
      End
    End If
End Sub

Private Sub fraCredits_DragDrop(Source As Control, X As Single, Y As Single)
    frmMain.fraCredits.Visible = False
    frmMain.tmrEnding.Enabled = False
    frmSplash.medMidi.URL = ""
    frmMain.mnuFile.Enabled = True
    frmMain.mnuOptions.Enabled = True
    frmMain.mnuHelp.Enabled = True
End Sub

Private Sub fraDefeat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Hell
    fraPaused.Visible = False
    'If you're in the middle of a game, warn before Restarting
        If booPlaying = True Then
            If MsgBox("This will restart the Level you 're currently at, talking away one life and subtacting 50 points.", vbOKCancel + vbCritical, "Marque Castle") = vbOK Then

                'Checks for a Game Over
                If intLivesNum > 0 Then
                    'Subtracts a Life
                    intLivesNum = intLivesNum - 1
                    frmMain.lblLives.Caption = intLivesNum
                    frmMain.tmrDeath.Enabled = True
                Else
                    MsgBox "Not enough lives to restart!" & vbNewLine & "Click ''File, New Game'' to begin again.", vbCritical, "Marque Castle"
                    Exit Sub
                End If

            Else
                Exit Sub
            End If
        End If
    fraDefeat.Visible = False
    frmSplash.medMidi.URL = ""
    If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
    strFile = lblFilePath.Caption
    intKeysNum = 0
    lblKeysNum.Caption = intKeysNum
    intCementBagsNum = 0
    lblCementBagsNum.Caption = intCementBagsNum
    LoadLevel
    booPlaying = True 'The game is in progress
    'Subtracts 50 Points from Score
        If dblScore > 50 Then
            dblScore = dblScore - 50
        Else
            dblScore = 0
        End If
        ScoreUpdate
    Exit Sub
Hell:
    MsgBox "Level data not found!" & vbNewLine & " Report this bug to marque@gavannon.com and tell us what happened!", vbCritical, "Marque Castle"
End Sub

Private Sub fraPaused_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If fraHighScore.Visible = True Or fraPictures.Visible = True Or fraSkinDir.Visible = True Then Exit Sub
    fraDefeat.Visible = False
    fraPaused.Visible = False
    mnuPauseGame.Caption = "&Pause Game"
     If fraHighScore.Visible = True Then Exit Sub
     If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
        'If there is a Drone Mouse, Enable him
             If booDroneMouse = True Then tmrDroneAI.Enabled = True
             If blnDeathMouse = True Then tmrDeathAI.Enabled = True
    tmrTimer.Enabled = True
End Sub

Private Sub imgHelp_Click(Index As Integer)
    If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
    If booPlaying = True And tmrTimer.Enabled = True Then
        fraPaused.Visible = True
        tmrDroneAI.Enabled = False
        tmrDeathAI.Enabled = False
        mnuPauseGame.Caption = "R&esume Game"
        tmrTimer.Enabled = False
    End If
    Select Case Index
        Case 0
            MsgBox "HELP!" & vbNewLine & vbNewLine & "The skins refer to the pictures that Marque Castle uses the actual game." & vbNewLine & "You can change the directory that it searches in when loading these ''Skins.''" & vbNewLine & vbNewLine & "  >> Make a copy the ''original'' skin directory, and replace" & vbNewLine & "      the image files with your own!  (Must have the same names).", vbInformation, "Marque Castle"
        Case 1
            MsgBox "HELP!" & vbNewLine & vbNewLine & "This area displays your current Items." & vbNewLine & "The first item is the number of Keys you have, and the second is the Bags of Cement you have." & vbNewLine & vbNewLine & "  >> Keys - Used to open doors (Your objective!)" & vbNewLine & "  >> Bags of Cement - Each Bag will turn water below you into Cement!  (Watch your quantity though!)" & vbNewLine & "  >> Metallic Boots - When you have these, you can walk freely on Spikes" & vbNewLine & "  >> Bombs - Used to blow up certain blocks within a 1 block radius (Can only carry 1 at a time!)" & vbNewLine & "  >> Clock - Resets the Time Remaining back to 150 (Can only carry 1 at a time!)", vbInformation, "Marque Castle"
        Case 2
            MsgBox "HELP!" & vbNewLine & vbNewLine & "Registration is no longer required for Marque Castle." & vbNewLine & "All the great features of the registered version are now free!" & vbNewLine & vbNewLine & "  What you'd get:" & vbNewLine & "    Access to the Scenario Creation Artist!" & vbNewLine & "      >> Able to play and create your own comple custom game in Marque Castle!" & vbNewLine & "      >> Change the skins for Marque Castle!" & vbNewLine & "      >> You can give Marque Castle your own, custom look!", vbInformation, "Marque Castle"
    End Select
End Sub

Private Sub lblCredits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.fraCredits.Visible = False
    frmMain.tmrEnding.Enabled = False
    frmSplash.medMidi.URL = ""
    frmMain.mnuFile.Enabled = True
    frmMain.mnuOptions.Enabled = True
    frmMain.mnuHelp.Enabled = True
End Sub

Private Sub lblDefeat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Hell
    fraPaused.Visible = False
    fraDefeat.Visible = False
    If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
    'If you're in the middle of a game, warn before Restarting
        If booPlaying = True Then If MsgBox("This will restart the Level you 're currently at, talking away one life and subtacting 50 points.", vbOKCancel + vbCritical, "Marque Castle") = vbCancel Then Exit Sub
    'Checks for a Game Over
        frmMain.lblLives.Caption = intLivesNum
        If intLivesNum = 0 Then
            MsgBox "Sorry dude, Game Over!", vbCritical, "Marque Castle"
            'Quits the Game
                QuitGame
            Exit Sub
        End If
    strFile = lblFilePath.Caption
    intKeysNum = 0
    lblKeysNum.Caption = intKeysNum
    intCementBagsNum = 0
    lblCementBagsNum.Caption = intCementBagsNum
    LoadLevel
    booPlaying = True 'The game is in progress
    'Subtracts 50 Points from Score
        If dblScore > 50 Then
            dblScore = dblScore - 50
        Else
            dblScore = 0
        End If
        ScoreUpdate
    Exit Sub
Hell:
    MsgBox "Level data not found!" & vbNewLine & " Report this bug to marque@gavannon.com and tell us what happened!", vbCritical, "Marque Castle"
End Sub

Private Sub lblPaused_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     If fraHighScore.Visible = True Or fraPictures.Visible = True Or fraSkinDir.Visible = True Then Exit Sub
    fraDefeat.Visible = False
    fraPaused.Visible = False
    mnuPauseGame.Caption = "&Pause Game"
     If fraHighScore.Visible = True Then Exit Sub
     If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
        'If there is a Drone Mouse, Enable him
             If booDroneMouse = True Then tmrDroneAI.Enabled = True
             If blnDeathMouse = True Then tmrDeathAI.Enabled = True
    tmrTimer.Enabled = True
End Sub
'
''  Begins Marque Castle Registration
'Private Sub lblRegistered_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'     'If the Registry Window is Closed
'          If fraRegistry.Visible = False Then
'               'If you're playing the Game
'                    If booPlaying = True And tmrTimer.Enabled = True Then
'                         'Pauses the Game
'                              fraPaused.Visible = True
'                              mnuPauseGame.Caption = "R&esume Game"
'                              fraPaused.Visible = True
'                              tmrTimer.Enabled = False
'                    End If
'               'Changes the Border Style for lblRegistered (The box you just Clicked)
'                    lblRegistered.BorderStyle = 1
'               'Shows the Registry Window
'                    fraRegistry.Visible = True
'     'If the Registry Window is Opened
'          Else
'               'Hides the Registry Window
'                    fraRegistry.Visible = False
'               'Changes the Border Style back to the original
'                    lblRegistered.BorderStyle = 0
'
'               'Resets the Values in the Registry Window
'                    txtEmailAddress.Text = ""
'                    txtFirstName.Text = ""
'                    txtLastName.Text = ""
'                    txtRegistryNumber.Text = ""
'                    chkRecieveEmail.Value = 1
'                    shpTotal.Visible = False
'                    shpPercent.Visible = False
'          End If
'End Sub
Private Sub lblRestore_Click()
    txtSkinDir.Text = "original"
End Sub

Private Sub lblSecret_Click(Index As Integer)
    If Index = 2 Then
        lblSecret(0).Visible = False
        lblSecret(1).Visible = False
        lblSecret(2).Visible = False
        If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
    End If
End Sub

Private Sub mnuAbout_Click()
  'Pauses the game
  If fraDefeat.Visible = False Then
    If booPlaying = True And tmrTimer.Enabled = True Then fraPaused.Visible = True
    mnuPauseGame.Caption = "R&esume Game"
  End If
  tmrTimer.Enabled = False
     
  'About message
  frmAbout.Show 1, Me
End Sub

Private Sub mnuAnimestudios_Click()
    GavannonCom.DoVerb
    If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
End Sub

Private Sub mnuBattleArena_Click()
     'Enables the Menu Commands
          mnuFile.Enabled = False
          mnuOptions.Enabled = False
          mnuHelp.Enabled = False

     'Pauses the Game if you're Playing
          If booPlaying = True And tmrTimer.Enabled = True Then

               'Pause the Game
                    fraPaused.Visible = True
                    mnuPauseGame.Caption = "R&esume Game"
                    tmrTimer.Enabled = False

          End If

     'Shows the Battle Arena Frame
          fraBattleArena.Visible = True

     'Message
          MsgBox "There are too many bugs for you to play this properly!" & vbNewLine & "You may check it out if you would like...", vbInformation, "Marque Castle"

End Sub

Private Sub mnuBestTimes_Click()

     'Pauses the game
        If fraDefeat.Visible = False Then
            If booPlaying = True And tmrTimer.Enabled = True Then fraPaused.Visible = True
            mnuPauseGame.Caption = "R&esume Game"
        End If
        tmrTimer.Enabled = False

     'Adjusts the Top Scores
          lblCongratulations.Caption = "These are the Best Times for this Level"
          cmdAction.Caption = "OK"
     
     'Shows the Best Times List
          fraHighScore.Visible = True

     'Disables the menus
         frmMain.mnuFile.Enabled = False
         frmMain.mnuOptions.Enabled = False
         frmMain.mnuHelp.Enabled = False

End Sub

Private Sub mnuHelpScenario_Click(Index As Integer)
    If frmMain.mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
    If booPlaying = True And tmrTimer.Enabled = True Then
        fraPaused.Visible = True
        tmrDroneAI.Enabled = False
        tmrDeathAI.Enabled = False
        mnuPauseGame.Caption = "R&esume Game"
        fraPaused.Visible = True
        tmrTimer.Enabled = False
    End If
    If Index = 0 Then
        MsgBox "HELP!" & vbNewLine & vbNewLine & "There are ''Rules'' as to what you can and can't save." & vbNewLine & vbNewLine & "  >> Every square must be filled" & vbNewLine & "  >> Keys: 1 or more" & vbNewLine & "  >> Doors: 1" & vbNewLine & "  >> George: No more than 1" & vbNewLine & "  >> Norman: 1" & vbNewLine & "  >> Toggle Blocks: No more than 250 (On or Off)", vbInformation, "Marque Castle"
    Else
        MsgBox "HELP!" & vbNewLine & vbNewLine & "The Level Creation Artist alows you to create your own Marque Castle game." & vbNewLine & "  >> Select one of the pictures on the left or right" & vbNewLine & "  >> Select where you want to place it on your map, by clicking that map area" & vbNewLine & "  >> Be sure you fill in all required info. for your level, displayed under ''More''", vbInformation, "Marque Castle"
    End If
End Sub

Private Sub mnuHelpSelect_Click(Index As Integer)
    If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
    If booPlaying = True And tmrTimer.Enabled = True Then
        fraPaused.Visible = True
        tmrDroneAI.Enabled = False
        tmrDeathAI.Enabled = False
        mnuPauseGame.Caption = "R&esume Game"
        fraPaused.Visible = True
        tmrTimer.Enabled = False
    End If
    Select Case Index
        Case 0
            MsgBox "HELP!" & vbNewLine & vbNewLine & "The skins refer to the pictures that Marque Castle uses the actual game." & vbNewLine & "You can change the directory that it searches in when loading these ''Skins.''" & vbNewLine & vbNewLine & "  >> Make a copy the ''original'' skin directory, and replace" & vbNewLine & "      the image files with your own!  (Must have the same names).", vbInformation, "Marque Castle"
        Case 1
            MsgBox "HELP!" & vbNewLine & vbNewLine & "This area displays your current Items." & vbNewLine & "The first item is the number of Keys you have, and the second is the Bags of Cement you have." & vbNewLine & vbNewLine & "  >> Keys - Used to open doors (Your objective!)" & vbNewLine & "  >> Bags of Cement - Each Bag will turn water below you into Cement!  (Watch your quantity though!)" & vbNewLine & "  >> metallic Boots - When you have these, you can walk freely on Spikes" & vbNewLine & "  >> Bombs - Used to blow up certain blocks within a 1 block radius (Can only carry 1 at a time!)" & vbNewLine & "  >> Clock - Resets the Time Remaining back to 150 (Can only carry 1 at a time!)", vbInformation, "Marque Castle"
        Case 2
            MsgBox "HELP!" & vbNewLine & vbNewLine & "Registration is no longer required for Marque Castle." & vbNewLine & "All the great features of the registered version are now free!" & vbNewLine & vbNewLine & "  What you'd get:" & vbNewLine & "    Access to the Scenario Creation Artist!" & vbNewLine & "      >> Able to play and create your own comple custom game in Marque Castle!" & vbNewLine & "      >> Change the skins for Marque Castle!" & vbNewLine & "      >> You can give Marque Castle your own, custom look!", vbInformation, "Marque Castle"
    End Select
End Sub

Private Sub mnuItemInfo_Click()
    If mnuItemInfo.Checked = True Then
        mnuItemInfo.Checked = False
        lblItemInfo.Visible = False
        tmrHideItemInfo.Enabled = False
    Else
        mnuItemInfo.Checked = True
        lblItemInfo.Caption = "Help will be displayed here throughout the game"
        lblItemInfo.Visible = True
        tmrHideItemInfo.Enabled = False
        tmrHideItemInfo.Enabled = True
    End If
End Sub

'Load Custom Game
Private Sub mnuLoadCustomGame_Click()

     'If Marque Castle is Registered
'          If strProperties(3) = "registered" Or lblRegistered.Visible = False Then

               'Adjusts the frmFile atributes acordinly
                    frmFile.Caption = "Data - Load"
                    frmFile.lblTitle.Caption = "Marque Castle"
                    frmFile.cmdSaveLoad.Caption = "Load"
                    frmFile.filFileListBox.Pattern = "*.cus"
                    frmFile.txtFileName.Text = "*.cus"
                    frmFile.lblFileName.Caption = "Scenario to Load:"
                    frmFile.lblFileName.Alignment = 2
                    frmFile.txtFileName.Visible = False

               'If you're Playing a Game
                    If booPlaying = True And tmrTimer.Enabled = True Then
                        fraPaused.Visible = True
                        mnuPauseGame.Caption = "R&esume Game"
                        tmrTimer.Enabled = False
                    End If

               'Show and Enable frmFile
                    frmFile.Show 1
                    frmFile.Enabled = True

               'Disable the Main Form
'                    frmMain.Enabled = False

'     'If Marque Castle isn't Registered
'          Else

'               'Displays a warning message
'                    MsgBox "Requires Registered version of Marque Castle." & vbNewLine & "You may Register by clicking ''Register Here.''", vbCritical, "Marque Castle"

'          End If

End Sub


'Toggles the Music On or Off
Private Sub mnuMusic_Click()
  
  'If Music is ON, make it OFF
  If mnuMusic.Checked = True Then
    'Stops any Music
    frmSplash.medMidi.URL = ""
    'Unchecks Music
    mnuMusic.Checked = False
  'If Music is OFF, make it ON
  Else
  
    'If you're playing the Game
    If booPlaying = True Then
      'Randomly Plays a MIDI Song
      frmSplash.medMidi.URL = App.Path & "\Music(" & CStr(RndBetween(0, 4)) & ").mid"
      'frmMain.Show
    End If
    
    'Checks Music
    mnuMusic.Checked = True
  End If
  
End Sub


Private Sub mnuNewGame_Click() 'Begins a new game
     
     'If in the middle of a game:
          If booPlaying = True Or fraMessage.Visible = True Or fraPaused.Visible = True Or fraDefeat.Visible = True Then

               'Warns you that this will start a new game before actually doing so
                    If MsgBox("This will Quit the current Game and begin a New one!", vbOKCancel, "Marque Castle") = vbCancel Then

                         'Leaves the Sub upon Cancelation
                              Exit Sub

                    End If

          End If

     'Resets the Score
          dblScore = 0
          ScoreUpdate

     'The Name of the File to Load:
          strFile = App.Path & "\Scenarios\LevelONE.lvl"

     'Loads the Level
          LoadLevel

     'Hides the Paused and Defeat Titles
          fraPaused.Visible = False
          fraDefeat.Visible = False

     'Loads the Pictures into their places
          imgGeorge.Picture = frmMain.pic(93).Picture
          imgKey.Picture = frmMain.pic(3).Picture
          imgCement.Picture = frmMain.pic(38).Picture
          imgBoots.Picture = frmMain.pic(35).Picture
          imgBoots.Visible = False
          imgClock.Picture = frmMain.pic(36).Picture
          imgClock.Visible = False
          imgBomb.Picture = frmMain.pic(37).Picture
          imgBomb.Visible = False
          imgExplosion.Picture = LoadPicture(strSkinDir & "\Explosion.gif")
          imgExplosionSmall.Picture = LoadPicture(strSkinDir & "\Explosion.gif")
          imgExplosion.Visible = False

     'Resets the Number of Lives
          If frmSplash.lblCheat.Visible = False Then
               intLivesNum = 5
          Else
               intLivesNum = 10
          End If
          lblLives.Caption = intLivesNum
          frmMain.tmrDeath.Enabled = True

End Sub


'Loads a Saved Game
Private Sub mnuLoadFile_Click()
  Dim strTempLives As String
  Dim strTempScore As String
  
    On Error GoTo Hell
    CommonDialog.Filter = "Marque Save Files (*.gav)|*.gav"
    CommonDialog.ShowOpen
    strGameSave = CommonDialog.FileName
    If strGameSave = "" Then Exit Sub
    Close #5
    Open strGameSave For Input As #5
    Input #5, strLevelFile, strTempLives, strTempScore
    lblLives.Caption = strTempLives
    intLivesNum = lblLives.Caption
    lblScore.Caption = strTempScore
    dblScore = CDbl(lblScore.Caption)
    strFile = strLevelFile
    LoadLevel
    fraPaused.Visible = False
    fraDefeat.Visible = False
     'Loads the Pictures into their places
          imgGeorge.Picture = frmMain.pic(93).Picture
          imgKey.Picture = frmMain.pic(3).Picture
          imgCement.Picture = frmMain.pic(38).Picture
          imgBoots.Picture = frmMain.pic(35).Picture
          imgClock.Picture = frmMain.pic(36).Picture
          imgBomb.Picture = frmMain.pic(37).Picture
          imgExplosion.Visible = False
          imgExplosion.Picture = LoadPicture(strSkinDir & "\Explosion.gif")
          imgExplosionSmall.Picture = LoadPicture(strSkinDir & "\Explosion.gif")
    'Hides the Boots, Clock, and the Bomb
        If frmMain.lblCheater.Visible = False Then frmMain.imgBoots.Visible = False
        frmMain.imgClock.Visible = False
        If frmMain.lblCheater.Visible = False Then frmMain.imgBomb.Visible = False
        If frmMain.lblCheater.Visible = True Then frmMain.imgBoots.Visible = True
        If frmMain.lblCheater.Visible = True Then frmMain.imgBomb.Visible = True
    Exit Sub
Hell:
    MsgBox "There was an unexpected error in loading your game!" & vbNewLine & "Contact marque@gavannon.com to report the error.", vbCritical, "Marque Castle"
End Sub


'Pauses the Game
Private Sub mnuPauseGame_Click()
     If fraHighScore.Visible = True Or fraPictures.Visible = True Or fraSkinDir.Visible = True Then Exit Sub

     If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"

     'If Unpaused, Pause
          If fraPaused.Visible = False Then

                fraPaused.Visible = True
                tmrDroneAI.Enabled = False
                tmrDeathAI.Enabled = False
                mnuPauseGame.Caption = "R&esume Game"
                fraPaused.Visible = True
                tmrTimer.Enabled = False

     'If Paused, Unpause
          Else

               'If there is a Drone Mouse, Enable him
                    If booDroneMouse = True Then tmrDroneAI.Enabled = True
                    If blnDeathMouse = True Then tmrDeathAI.Enabled = True
               
               fraPaused.Visible = False
               mnuPauseGame.Caption = "&Pause Game"
               fraPaused.Visible = False
               tmrTimer.Enabled = True

     End If
End Sub


'Quits the Game
Private Sub mnuQuitGame_Click()

    'Pause
        fraPaused.Visible = True
        tmrDroneAI.Enabled = False
        tmrDeathAI.Enabled = False
        mnuPauseGame.Caption = "R&esume Game"
        fraPaused.Visible = True
        tmrTimer.Enabled = False

     If mnuSound.Checked = True Then
          'Makes a beeping sound
               PlaySound 0, App.Path & "\Beep.wav"
     End If

     If MsgBox("This will Quit your current Game.", vbOKCancel, "Marque Castle") = vbCancel Then

          'Exits Sub upon Cancelation
               Exit Sub

     End If

     'Quits the Game
          QuitGame

End Sub


'Loads the Scenario Creation Artist
Private Sub mnuCreateCustomGame_Click()

'    If Mid(strProperties(3), 1, 2) <> "re" Or lblRegistered.Visible = True Then
'        MsgBox "Requires Registered version of Marque Castle." & vbNewLine & "You may Register by clicking ''Register Here.''", vbCritical, "Marque Castle"
'        Exit Sub
'    End If

    'If you are playing a Game
    If booPlaying = True Or fraPaused.Visible = True Or fraDefeat.Visible = True Or fraMessage.Visible = True Or cmdBegin02.Visible = True Then

        'Pause
        If fraHighScore.Visible = False And fraPictures.Visible = False And fraSkinDir.Visible = False And fraDefeat.Visible = False Then
            fraPaused.Visible = True
            tmrDroneAI.Enabled = False
            tmrDeathAI.Enabled = False
            mnuPauseGame.Caption = "R&esume Game"
            fraPaused.Visible = True
            tmrTimer.Enabled = False
        End If

        'Asks if you would like to quit first
        If MsgBox("You must Quit your current Game first!", vbOKCancel, "Marque Castle") = vbOK Then

              If mnuSound.Checked = True Then
                   'Makes a beeping sound
                        PlaySound 0, App.Path & "\Beep.wav"
              End If
              'Quits the Game
              QuitGame
        
        Else

            Exit Sub

        End If
    
    End If

    'Show and Enable the Scenario Creation Artist
         frmCreate.Show 1
         frmCreate.Enabled = True
End Sub


'Exits the Game
Private Sub mnuExit_Click()
  Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

     'If a Game is in progress and Unpaused
          If booPlaying = True And fraPaused.Visible = False And fraDefeat.Visible = False Then
          If lblEnd(0).Visible = True Then Exit Sub
               'What Key was pressed
                    Select Case KeyCode
'_____________________________________________________________________________________________
                         'If you've pressed <UP>
                              Case Is = vbKeyUp

                                   'Checks World Border
                                        If intPos > 19 Then

                                             'Moves George to the appropriate possition
                                            frmMain.imgMap(intPos).Picture = frmMain.pic(91).Picture
                                            MoveGeorge -20, 91, True

                                        End If
'_____________________________________________________________________________________________
                         'If you've pressed <W>
                              Case Is = vbKeyW

                                   'If there is no Norman
                                        If intNormanPosition < 0 Then

                                             Exit Sub

                                        End If

                                   'Checks World Border
                                        If intNormanPosition > 19 Then

                                             'Moves George to the appropriate possition
                                            frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(891).Picture
                                            MoveGeorge -20, 891, False

                                        End If
'_____________________________________________________________________________________________
                         'If you've pressed <DOWN>
                              Case Is = vbKeyDown

                                   'Checks World Border
                                        If intPos < 380 Then

                                             'Moves George to the appropriate possition
                                            frmMain.imgMap(intPos).Picture = frmMain.pic(93).Picture
                                            MoveGeorge 20, 93, True

                                        End If
'_____________________________________________________________________________________________
                         'If you've pressed <S>
                              Case Is = vbKeyS

                                   'If there is no Norman
                                        If intNormanPosition < 0 Then

                                             Exit Sub

                                        End If

                                   'Checks World Border
                                        If intNormanPosition < 380 Then

                                             'Moves George to the appropriate possition
                                            frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(893).Picture
                                            MoveGeorge 20, 893, False

                                        End If
'_____________________________________________________________________________________________
                         'If you've pressed <RIGHT>
                              Case Is = vbKeyRight

                                   'Checks World Border
                                        If intPos Mod 20 <> 19 Then

                                             'Moves George to the appropriate possition
                                            frmMain.imgMap(intPos).Picture = frmMain.pic(92).Picture
                                            MoveGeorge 1, 92, True

                                        End If
'_____________________________________________________________________________________________
                         'If you've pressed <D>
                              Case Is = vbKeyD

                                   'If there is no Norman
                                        If intNormanPosition < 0 Then

                                             Exit Sub

                                        End If

                                   'Checks World Border
                                        If intNormanPosition Mod 20 <> 19 Then

                                             'Moves George to the appropriate possition
                                            frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(892).Picture
                                            MoveGeorge 1, 892, False

                                        End If
'_____________________________________________________________________________________________
                         'If you've pressed <LEFT>
                              Case Is = vbKeyLeft

                                   'Checks World Border
                                        If intPos Mod 20 <> 0 Then

                                             'Moves George to the appropriate possition
                                              frmMain.imgMap(intPos).Picture = frmMain.pic(94).Picture
                                              MoveGeorge -1, 94, True

                                        End If
'_____________________________________________________________________________________________
                         'If you've pressed <A>
                              Case Is = vbKeyA

                                   'If there is no Norman
                                        If intNormanPosition < 0 Then

                                             Exit Sub

                                        End If

                                   'Checks World Border
                                        If intNormanPosition Mod 20 <> 0 Then

                                             'Moves George to the appropriate possition
                                            frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(894).Picture
                                            MoveGeorge -1, 894, False

                                        End If
'_____________________________________________________________________________________________
                        'If you've pressed <Space Bar>
                                Case Is = vbKeySpace

                                    'Checks to see if you have a Bomb
                                    If imgBomb.Visible = True Then
     
                                      'Removes the Bomb pic
                                      If lblCheater.Visible = False Then imgBomb.Visible = False
                                                
                                      'Makes an explosion sound
                                      If mnuSound.Checked = True Then PlaySound 0, App.Path & "\UseBomb.wav"

                                      'Makes the Explosion
                                      Explosion intPos
                                    End If
                              End Select

     'If no Game is in progress
     ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
          If fraPaused.Visible = True Then MsgBox "To continue your game, click 'Options', 'Resume Game', or press F12.", vbInformation, "Marque Castle"
     End If

    If lblSpeed.Visible = True Then
        If KeyCode = 109 Or KeyCode = 189 Then               ' - speed
            If tmrDroneAI.Interval <= 2000 Then tmrDroneAI.Interval = tmrDroneAI.Interval + 20
            If tmrDeathAI.Interval <= 2000 Then tmrDeathAI.Interval = tmrDeathAI.Interval + 20
        ElseIf KeyCode = 107 Or KeyCode = 187 Then        ' + speed
            If tmrDroneAI.Interval >= 21 Then tmrDroneAI.Interval = tmrDroneAI.Interval - 20
            If tmrDeathAI.Interval >= 21 Then tmrDeathAI.Interval = tmrDeathAI.Interval - 20
        End If
    End If

End Sub

Private Sub mnuRestartLevel_Click()
    On Error GoTo Hell
    fraPaused.Visible = False
    'If you're in the middle of a game, warn before Restarting
        If booPlaying = True Then
            If MsgBox("This will restart the Level you 're currently at, talking away one life and subtacting 50 points.", vbOKCancel + vbCritical, "Marque Castle") = vbOK Then

                'Checks for a Game Over
                If intLivesNum > 0 Then
                    'Subtracts a Life
                    intLivesNum = intLivesNum - 1
                    frmMain.lblLives.Caption = intLivesNum
                    frmMain.tmrDeath.Enabled = True
                Else
                    MsgBox "Not enough lives to restart!" & vbNewLine & "Click ''File, New Game'' to begin again.", vbCritical, "Marque Castle"
                    frmMain.tmrDeath.Enabled = True
                    Exit Sub
                End If

            Else
                Exit Sub
            End If
        End If
    fraDefeat.Visible = False
    frmSplash.medMidi.URL = ""
    If mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
    strFile = lblFilePath.Caption
    intKeysNum = 0
    lblKeysNum.Caption = intKeysNum
    intCementBagsNum = 0
    lblCementBagsNum.Caption = intCementBagsNum
    LoadLevel
    booPlaying = True 'The game is in progress
    'Subtracts 50 Points from Score
        If dblScore > 50 Then
            dblScore = dblScore - 50
        Else
            dblScore = 0
        End If
        ScoreUpdate
    Exit Sub
Hell:
    MsgBox "Level data not found!" & vbNewLine & " Report this bug to marque@gavannon.com and tell us what happened!", vbCritical, "Marque Castle"
End Sub

Private Sub mnuSaveGame_Click()
    SaveGame
End Sub

Private Sub mnuSaveGameAs_Click()
    CommonDialog.Filter = "Marque Save Files (*.gav)|*.gav"
    CommonDialog.ShowSave
    strGameSave = CommonDialog.FileName
    If strGameSave = "" Then Exit Sub
    If Mid$(strGameSave, Len(strGameSave) - 3, 4) <> ".gav" Then strGameSave = strGameSave & ".gav"
    SaveGame
End Sub

'  Changes the Skin Directory
Private Sub mnuSkins_Click()

'          If Mid$(strProperties(3), 1, 2) <> "re" Or lblRegistered.Visible = True Then
'            MsgBox "Requires Registered version of Marque Castle." & vbNewLine & "You may Register by clicking ''Register Here.''", vbCritical, "Marque Castle"
'            Exit Sub
'          End If

     'Pauses the Game if you're Playing
          If booPlaying = True And tmrTimer.Enabled = True Then

               'Pause the Game
                    fraPaused.Visible = True
                    mnuPauseGame.Caption = "R&esume Game"
                    fraPaused.Visible = True
                    tmrTimer.Enabled = False

          End If

     'Shows the Skin Directory Frame
          fraSkinDir.Visible = True

     'Shows the Pictures
          fraPictures.Visible = True

     'Disables all the Menus
          mnuFile.Enabled = False
          mnuOptions.Enabled = False
          mnuHelp.Enabled = False

End Sub


Private Sub mnuSound_Click()

     'If Sound is ON, make it OFF
          If mnuSound.Checked = True Then

               'Unchecks Sound
                    mnuSound.Checked = False

     'If Sound is OFF, make it ON
          Else

               'Makes an beeping sound
                    PlaySound 0, App.Path & "\Beep.wav"

               'Checks Sound
                    mnuSound.Checked = True

          End If

End Sub

Private Sub optMultiLevel_GotFocus(Index As Integer)
    cmdStartMulti.Default = True
End Sub

Private Sub SaveGame()
    On Error GoTo Hell
    If strGameSave = "" Then
        CommonDialog.Filter = "Marque Save Files (*.gav)|*.gav"
        CommonDialog.ShowSave
        strGameSave = CommonDialog.FileName
        If strGameSave = "" Then Exit Sub
        If Mid$(strGameSave, Len(strGameSave) - 3, 4) <> ".gav" Then strGameSave = strGameSave & ".gav"
    End If
    Close #5
    Open strGameSave For Output As #5
    Write #5, strLevelFile, lblLives, lblScore
    Exit Sub
Hell:
    MsgBox "There was an unexpected error in saving your game!" & vbNewLine & "Contact marque@gavannon.com to report the error.", vbCritical, "Marque Castle"
End Sub


Private Sub tmrCement_Timer()
  Static CementShown As Byte

  If CementShown <= 7 Then
    CementShown = CementShown + 1
    lblCementBagsNum.Visible = IIf(lblCementBagsNum.Visible, False, True)
  Else
    tmrCement.Enabled = False
    CementShown = 0
    lblCementBagsNum.Visible = True
  End If
End Sub

Private Sub tmrDeath_Timer()
  Static DeathBlink As Byte

  If DeathBlink <= 15 Then
    DeathBlink = DeathBlink + 1
    lblLives.Visible = IIf(lblLives.Visible, False, True)
  Else
    tmrDeath.Enabled = False
    DeathBlink = 0
    lblLives.Visible = True
  End If

End Sub

'Adversiary that moves randomly
Private Sub tmrDeathAI_Timer()
  Dim z As Integer

    If lblEnd(0).Visible = True Or fraMessage.Visible = True Or cmdBegin02.Visible = True Then
        tmrDroneAI.Enabled = False
        tmrDeathAI.Enabled = False
        Exit Sub
    End If
    On Error Resume Next
    If fraPaused.Visible = True Or fraMessage.Visible = True Or intDeathPos < 0 Then Exit Sub
    If blnDeathMouse = False Then
        tmrDeathAI.Enabled = False
        Exit Sub
    End If
    
    'Randomly adjust its speed
    tmrDeathAI.Interval = RndBetween(1, 150) + IIf(intDeathPos > intPos, 100, 0)
  
  
    '1 in 3 chance to move towards you
    z = RndBetween(1, 3)
    
    'Move towards Norm
    If z = 3 Then
      If intDeathPos + 20 > intPos Then
        z = 1
      ElseIf intDeathPos - 20 < intPos Then
        z = 2
      End If
    Else
      z = RndBetween(1, 4)
    End If
    
    Select Case z
        Case 1      ' Up
            If ((intDeathPos - 20) > 19 And CInt(strCell(intDeathPos - 20)) < 2) Or ((intDeathPos - 20) > 19 And CInt(strCell(intDeathPos - 20)) = intPos) Then
                If intDeathPos > 19 Then
                    imgMap(intDeathPos).Picture = pic(CInt(strDeathGround)).Picture
                    strCell(intDeathPos) = CInt(strDeathGround)
                    intDeathPos = intDeathPos - 20
                    strDeathGround = strCell(intDeathPos)
                    imgMap(intDeathPos).Picture = pic(992).Picture
                    strCell(intDeathPos) = "992"
                End If
            End If
        Case 2      ' Down
            If ((intDeathPos + 20) < 380 And CInt(strCell(intDeathPos + 20)) < 2) Or ((intDeathPos + 20) < 380 And CInt(strCell(intDeathPos + 20)) = intPos) Then
                If intDeathPos < 380 Then
                    imgMap(intDeathPos).Picture = pic(CInt(strDeathGround)).Picture
                    strCell(intDeathPos) = CInt(strDeathGround)
                    intDeathPos = intDeathPos + 20
                    strDeathGround = strCell(intDeathPos)
                    imgMap(intDeathPos).Picture = pic(992).Picture
                    strCell(intDeathPos) = "992"
                End If
            End If
        Case 3      ' Left
            If ((intDeathPos - 1) Mod 20 <> 0 And CInt(strCell(intDeathPos - 1)) < 2) Or ((intDeathPos + 20) < 380 And CInt(strCell(intDeathPos - 1)) = intPos) Then
                If intDeathPos Mod 20 <> 0 Then
                    imgMap(intDeathPos).Picture = pic(CInt(strDeathGround)).Picture
                    strCell(intDeathPos) = CInt(strDeathGround)
                    intDeathPos = intDeathPos - 1
                    strDeathGround = strCell(intDeathPos)
                    imgMap(intDeathPos).Picture = pic(992).Picture
                    strCell(intDeathPos) = "992"
                End If
            End If
        Case 4      ' Right
            If ((intDeathPos + 1) Mod 20 <> 19 And CInt(strCell(intDeathPos + 1)) < 2) Or ((intDeathPos + 20) < 380 And CInt(strCell(intDeathPos + 1)) = intPos) Then
                If intDeathPos Mod 20 <> 19 Then
                    imgMap(intDeathPos).Picture = pic(CInt(strDeathGround)).Picture
                    strCell(intDeathPos) = CInt(strDeathGround)
                    intDeathPos = intDeathPos + 1
                    strDeathGround = strCell(intDeathPos)
                    imgMap(intDeathPos).Picture = pic(992).Picture
                    strCell(intDeathPos) = "992"
                End If
            End If
    End Select
    'TempGround = strCell(intPos)
    
    'George moved onto it
         If intDronePos = intPos Then

              'Defeat
                   If frmMain.mnuSound.Checked = True Then
                        'Makes a beeping sound
                             PlaySound 0, App.Path & "\Defeat.wav"
                   End If

                   Defeat

                 If frmMain.mnuItemInfo.Checked = True Then
                     frmMain.lblItemInfo.Caption = "Watch out for Mice!"
                     frmMain.lblItemInfo.Visible = True
                     frmMain.tmrHideItemInfo.Enabled = False
                     frmMain.tmrHideItemInfo.Enabled = True
                 End If

         End If
End Sub


'Adversiary that follows the Walls
Private Sub tmrDroneAI_Timer()

    If lblEnd(0).Visible = True Or fraMessage.Visible = True Or cmdBegin02.Visible = True Then
        tmrDroneAI.Enabled = False
        tmrDeathAI.Enabled = False
        Exit Sub
    End If
     'If the Game is Paused
          If fraPaused.Visible = True Or fraMessage.Visible = True Or intDronePos < 0 Then Exit Sub

     'If there is a live Drone Mouse on the Map
          If booDroneMouse = True Then

               'If it couldn't move for 3 turns
                    If intDroneUnmove = 3 Then

                         'Changes the picture to the Dead Drone Mouse
                              imgMap(intDronePos).Picture = frmMain.pic(93).Picture

               'If it couldn't move for 4 turns
                    ElseIf intDroneUnmove = 4 Then

                         'Changes the picture to the Dead Drone Mouse
                              imgMap(intDronePos).Picture = frmMain.pic(CInt(strDefaultGround)).Picture

                         'Sets the Grid Container accordingly
                              strCell(intPos) = strGround

               'If it could move
                    Else

                         'Moves the Drone Mouse one square
                          DroneMouseAI

                    End If

     'If there isn't a Live Drone Mouse on the Map
          Else

               'Stops the Drone Mouse AI Timer
                    tmrDroneAI.Enabled = False

          End If

End Sub


Private Sub tmrEnding_Timer()
  Dim i As Integer
  
    Select Case intCreditNum
        Case 0
            lblCredits.Caption = "Marque Castle" & vbNewLine & "Version 1.2"
        Case 1
            lblCredits.Caption = "Programmed by" & vbNewLine & "Chris Ringrose"
        Case 2
            lblCredits.Caption = "Game design" & vbNewLine & "Chris Ringrose"
        Case 3
            lblCredits.Caption = "Producer" & vbNewLine & "Chris Ringrose"
        Case 4
            lblCredits.Caption = "Level design by" & vbNewLine & "Chris Ringrose"
        Case 5
            lblCredits.Caption = "Story and dialog" & vbNewLine & "Chris Ringrose"
        Case 6
            lblCredits.Caption = "Quality assurance" & vbNewLine & "Chris Ringrose"
        Case 7
            lblCredits.Caption = "Game testers" & vbNewLine & "Chris, Wallace, Justin, Mikko, Neal, and Eric!"
        Case 8
            lblCredits.Caption = "''Defeat'' and ''water'' sounds by" & vbNewLine & "Chris Ringrose"
        Case 9
            lblCredits.Caption = "All other sounds from" & vbNewLine & "''Resident Evil''"
        Case 10
            lblCredits.Caption = "End music from Nintendo's" & vbNewLine & "''EarthBound''"
        Case 11
            lblCredits.Caption = "Marque Castle theme song" & vbNewLine & "from ... unknown!  (Let me know!)"
        Case 12
            lblCredits.Caption = "All the other music came with" & vbNewLine & "Yamaha's XGPlayer"
        Case 13
            lblCredits.Caption = vbNewLine & "Special Thanks to..."
        Case 14
            lblCredits.Caption = "Neal, Wallace, Justin, Nick, Mikko, Eric, Pepsi (for the late nights)..."
        Case 15
            lblCredits.Caption = "...Pepsi (for the early mornings)," & vbNewLine & "and SARAH!  (Hi babes! *french*)"
        Case 16
            lblCredits.Caption = "Everything else by" & vbNewLine & "Chris Ringrose"
        Case 17
            lblCredits.Caption = "Copyright © 2003 Chris Ringrose" & vbNewLine & "marque@gavannon.com"
        Case 18
            lblCredits.Caption = "http://www.gavannon.com/" & vbNewLine & "http://www.planetsourcecode.com/" & vbNewLine & "http://www.sourcecode4free.com/"
        Case 19
            lblCredits.Caption = ""
        Case 20
            lblCredits.Caption = vbNewLine & "The End"
        Case 21
            lblCredits.Caption = ""
        Case 22
            lblCredits.Caption = ""
        Case 23
            lblCredits.Caption = vbNewLine & "Look, there's nothing more, okay?!"
        Case 24
            lblCredits.Caption = ""
        Case 25
            lblCredits.Caption = ""
        Case 26
            lblCredits.Caption = ""
        Case 27
            lblCredits.Caption = vbNewLine & "I'll delete your Windows directory!"
        Case 28
            lblCredits.Caption = ""
        Case 29
            lblCredits.Caption = ""
        Case 30
            lblCredits.Caption = vbNewLine & "Really, I will.  It isn't hard to do ..."
        Case 31
            lblCredits.Caption = ""
        Case 32
            lblCredits.Caption = ""
        Case 33
            lblCredits.Caption = "All I need is an API to locate it..."
        Case 34
            lblCredits.Caption = vbNewLine & "And the VB Kill command ..."
        Case 35
            lblCredits.Caption = ""
        Case 36
            lblCredits.Caption = ""
        Case 37
            lblCredits.Caption = vbNewLine & "Reformating Hard Drive in ..."
        Case 38
            lblCredits.Caption = vbNewLine & "5"
        Case 39
            lblCredits.Caption = vbNewLine & "4"
        Case 40
            lblCredits.Caption = vbNewLine & "3"
        Case 41
            lblCredits.Caption = vbNewLine & "2"
        Case 42
            lblCredits.Caption = vbNewLine & "1"
        Case 43
            lblCredits.Caption = ""
        Case 44
            fraCredits.Visible = False
            tmrEnding.Enabled = False
            frmSplash.medMidi.URL = ""
            mnuFile.Enabled = True
            mnuOptions.Enabled = True
            mnuHelp.Enabled = True
    End Select
    For i = 0 To 255
        lblCredits.ForeColor = RGB(i, i, i)
        Sleep 10
        DoEvents
    Next
    Sleep 2000
    Do While i > 0
        lblCredits.ForeColor = RGB(i, i, i)
        Sleep 2
        DoEvents
        i = i - 1
    Loop
    intCreditNum = intCreditNum + 1
End Sub

Private Sub tmrExplosion_Timer()
  'Destroys Blocks
  imgExplosion.Visible = False
  tmrExplosion.Enabled = False
End Sub


Private Sub tmrExplosionSmall_Timer()
  imgExplosionSmall.Visible = IIf(imgExplosionSmall.Visible, False, True)
  If imgExplosionSmall.Visible = False Then tmrExplosionSmall.Enabled = False
End Sub

Private Sub tmrHideItemInfo_Timer()
    lblItemInfo.Visible = False
    tmrHideItemInfo.Enabled = False
End Sub

Private Sub tmrKey_Timer()
  Static KeyShown As Byte

  If KeyShown <= 7 Then
    KeyShown = KeyShown + 1
    lblKeysNum.Visible = IIf(lblKeysNum.Visible, False, True)
  Else
    tmrKey.Enabled = False
    KeyShown = 0
    lblKeysNum.Visible = True
  End If
End Sub

Private Sub tmrNewGame_Timer()
  If lblItemInfo.Caption = "" Then
    lblItemInfo.Caption = "To start a new game, click File, New Game!"
  Else
    lblItemInfo.Caption = ""
  End If
End Sub

Private Sub tmrPts_Timer()

     'Hides the Points
          lblPts.Visible = False

     'Disables the Timer
          tmrPts.Enabled = False

End Sub


'  Changes the colour of the score numbers back
Private Sub tmrScore_Timer()

     'Changes the colour back to Black
          lblScore.ForeColor = &H0&
          lblSteps(1).ForeColor = &H0&

     'Disables the Timer
          tmrScore.Enabled = False

End Sub


'The Time Limit durring run time
Private Sub tmrTimer_Timer()
  'Dim Direction As Integer
  
    If lblEnd(0).Visible = True Then
        tmrTimer.Enabled = False
        Exit Sub
    End If
     
     If fraPaused.Visible = False Then

          If CInt(lblTimer.Caption) > 0 Then

              lblTimer.Caption = CStr(CInt(lblTimer.Caption) - 1)
              If CInt(lblTimer.Caption) <= 15 Then
                Beep
                tmrWatchTime.Enabled = True
              End If

          Else

                    'You are no longer playing (You Died)
                         booPlaying = False
                         intLivesNum = intLivesNum - 1
                         lblLives.Caption = intLivesNum
                         tmrWatchTime.Enabled = True
                         tmrDeath.Enabled = True

                    'Checks for a Game Over
                         If intLivesNum < 1 Then

                              lblLives.Caption = intLivesNum

                         End If

                    'Removes George (From Original Spot)
                         PicBack

                    'Places the new Possition and Picture (Adding George)
'                         intPos = intPos + Direction
                         imgMap(intPos).Picture = frmMain.pic(95).Picture

                    If mnuSound.Checked = True Then
                         'Makes an defeat sound
                              PlaySound 0, App.Path & "\Defeat.wav"
                    End If

                    Defeat

               End If

     End If

End Sub

Private Sub tmrWatchTime_Timer()
  Static Blink As Byte
  
  If Blink <= 10 Then
    lblTimer.Visible = IIf(lblTimer.Visible, False, True)
    Blink = Blink + 1
  Else
    tmrWatchTime.Enabled = False
    tmrWatchTime.Interval = 350
    Blink = 0
    lblTimer.Visible = True
  End If
End Sub

Private Sub txtCurrentScore_Change()
    If Len(txtCurrentScore.Text) = 0 Then
        cmdAction.Enabled = False
    Else
        cmdAction.Enabled = True
    End If
End Sub

Private Sub txtCurrentScore_GotFocus()
    cmdAction.Default = True
End Sub

Private Sub txtEmailAddress_GotFocus()
    cmdRegister.Default = True
End Sub

Private Sub txtFirstName_GotFocus()
    cmdRegister.Default = True
End Sub

Private Sub txtMessage_GotFocus()
    cmdBegin.Default = True
End Sub

Private Sub txtSkinDir_Change()
    If Len(txtSkinDir) = 0 Then
        cmdApply.Enabled = False
    Else
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtSkinDir_GotFocus()
    cmdApply.Default = True
End Sub

