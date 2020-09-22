VERSION 5.00
Begin VB.Form frmCreate 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scenario Creation Artist - Untitled"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
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
   Icon            =   "frmCreate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "More"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6420
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Toggle Level Manager (Shows or hides)"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Frame fraLevelManager 
      BackColor       =   &H000080FF&
      Caption         =   "Level Manager:"
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
      Height          =   540
      Left            =   142
      TabIndex        =   21
      Top             =   5595
      Width           =   6615
      Begin VB.TextBox txtFilePath 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   225
         Width           =   4395
      End
      Begin VB.Frame fraBestTimes 
         BackColor       =   &H000080FF&
         Caption         =   "Best Time:"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   540
         TabIndex        =   34
         Top             =   2520
         Width           =   5475
         Begin VB.Label lblBestTimesDir 
            BackColor       =   &H0080C0FF&
            Caption         =   "Untitled Level.bt"
            Height          =   195
            Left            =   2700
            TabIndex        =   36
            ToolTipText     =   "The Best Times File Name"
            Top             =   240
            Width           =   2265
         End
         Begin VB.Label lblBestTimes 
            BackColor       =   &H000080FF&
            Caption         =   "File where Best Times are stored:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   600
            TabIndex        =   35
            ToolTipText     =   "Loads after victory"
            Top             =   240
            Width           =   2115
         End
      End
      Begin VB.TextBox txtAuthor 
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4650
         MaxLength       =   35
         TabIndex        =   8
         ToolTipText     =   "The Author of this Level"
         Top             =   2100
         Width           =   1815
      End
      Begin VB.TextBox txtLevelTitle 
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   975
         MaxLength       =   25
         TabIndex        =   3
         Text            =   "Untitled Level"
         ToolTipText     =   "Title of Level (Displayed to Gamer)"
         Top             =   750
         Width           =   2265
      End
      Begin VB.Frame fraLevelOrdering 
         BackColor       =   &H000080FF&
         Caption         =   "Level Ordering:"
         ForeColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   225
         TabIndex        =   24
         Top             =   1125
         Width           =   3540
         Begin VB.ComboBox cboLevelNumber 
            BackColor       =   &H0080C0FF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            ItemData        =   "frmCreate.frx":030A
            Left            =   1200
            List            =   "frmCreate.frx":03A4
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "The Level Number (1 to 50)"
            Top             =   225
            Width           =   915
         End
         Begin VB.CheckBox chkLastLevel 
            BackColor       =   &H000080FF&
            Caption         =   "Last Level"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   225
            TabIndex        =   5
            ToolTipText     =   "No more levels"
            Top             =   600
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.TextBox txtNextLevel 
            BackColor       =   &H0080C0FF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1050
            TabIndex        =   6
            Text            =   "*.cus"
            ToolTipText     =   "Loads after victory"
            Top             =   900
            Width           =   2265
         End
         Begin VB.Label lblLastLevelInfo 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "Must be in same directory"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1050
            TabIndex        =   31
            Top             =   675
            Width           =   2265
         End
         Begin VB.Label lblLevelNumber 
            BackColor       =   &H000080FF&
            Caption         =   "Level Number:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   225
            TabIndex        =   27
            ToolTipText     =   "The Level Number (1 to 50)"
            Top             =   225
            Width           =   990
         End
         Begin VB.Label lblNextLevel 
            BackColor       =   &H000080FF&
            Caption         =   "Next Level:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   225
            TabIndex        =   25
            ToolTipText     =   "Loads after victory"
            Top             =   900
            Width           =   765
         End
      End
      Begin VB.TextBox txtMessage 
         BackColor       =   &H0080C0FF&
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
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   3975
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmCreate.frx":0467
         ToolTipText     =   "Displayed to Gamer on Startup"
         Top             =   975
         Width           =   2490
      End
      Begin VB.CommandButton cmdLoad 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Load specified File"
         Top             =   225
         Width           =   765
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4575
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Save to specified File"
         Top             =   225
         Width           =   765
      End
      Begin VB.Label lblAuthor 
         BackColor       =   &H000080FF&
         Caption         =   "Author:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3975
         TabIndex        =   30
         ToolTipText     =   "The Author of this Level"
         Top             =   2100
         Width           =   540
      End
      Begin VB.Label lblLevelTitle 
         BackColor       =   &H000080FF&
         Caption         =   "Level Title:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   225
         TabIndex        =   26
         ToolTipText     =   "Title of Level (Displayed to Gamer)"
         Top             =   750
         Width           =   765
      End
      Begin VB.Label lblMessage 
         BackColor       =   &H000080FF&
         Caption         =   "Message to Gamer on Startup:"
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   3975
         TabIndex        =   23
         ToolTipText     =   "Displayed to Gamer on Startup"
         Top             =   750
         Width           =   1965
      End
   End
   Begin VB.CommandButton cmdUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Undo"
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
      Height          =   285
      Left            =   5902
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4365
      Width           =   1020
   End
   Begin VB.CheckBox chkPaintFill 
      BackColor       =   &H000080FF&
      Caption         =   "Paint Fill"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   52
      TabIndex        =   39
      Top             =   4365
      Width           =   1050
   End
   Begin VB.Frame fraError 
      BackColor       =   &H000080FF&
      Caption         =   "Error in Saving!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1830
      Left            =   1492
      TabIndex        =   18
      Top             =   225
      Visible         =   0   'False
      Width           =   3915
      Begin VB.ComboBox cboErrors 
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   75
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Error Listing"
         Top             =   1065
         Width           =   3765
      End
      Begin VB.CommandButton cmdResume 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Resume"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2055
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Resume to fix Errors"
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label lblErrorDescript 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "There was one or more errors found in your custom level, and saving did not commence."
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
         Height          =   480
         Left            =   75
         TabIndex        =   37
         Top             =   375
         Width           =   3765
      End
      Begin VB.Image imgHelp 
         Height          =   330
         Index           =   2
         Left            =   3030
         Picture         =   "frmCreate.frx":0486
         ToolTipText     =   "Help on Registration"
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "C o u l d   n o t   S a v e !"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Error in Saving"
         Top             =   225
         Width           =   3765
      End
      Begin VB.Label lblDetails 
         BackColor       =   &H000080FF&
         Caption         =   "Error(s):"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   75
         TabIndex        =   29
         Top             =   885
         Width           =   690
      End
   End
   Begin VB.Frame fraToggleBlocks 
      BackColor       =   &H000080FF&
      Caption         =   "Toggle Blocks:"
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
      Height          =   645
      Left            =   5842
      TabIndex        =   17
      Top             =   3630
      Width           =   1155
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   33
         Left            =   600
         Picture         =   "frmCreate.frx":084A
         Stretch         =   -1  'True
         ToolTipText     =   "On (You can Walk On)"
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   34
         Left            =   300
         Picture         =   "frmCreate.frx":0B8C
         Stretch         =   -1  'True
         ToolTipText     =   "Off (You can't Walk On)"
         Top             =   255
         Width           =   240
      End
   End
   Begin VB.Frame fraFountains 
      BackColor       =   &H000080FF&
      Caption         =   "Fountains:"
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
      Height          =   1890
      Left            =   5962
      TabIndex        =   16
      Top             =   1680
      Width           =   840
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   32
         Left            =   450
         Picture         =   "frmCreate.frx":0ECE
         Stretch         =   -1  'True
         ToolTipText     =   "Bottom Right"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   225
         Index           =   24
         Left            =   150
         Picture         =   "frmCreate.frx":1210
         Stretch         =   -1  'True
         ToolTipText     =   "Top Left"
         Top             =   300
         Width           =   225
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   26
         Left            =   450
         Picture         =   "frmCreate.frx":1552
         Stretch         =   -1  'True
         ToolTipText     =   "Top Right"
         Top             =   300
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   28
         Left            =   150
         Picture         =   "frmCreate.frx":1894
         Stretch         =   -1  'True
         ToolTipText     =   "Water in Fountain"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   30
         Left            =   150
         Picture         =   "frmCreate.frx":1BD6
         Stretch         =   -1  'True
         ToolTipText     =   "Bottom Left"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   25
         Left            =   150
         Picture         =   "frmCreate.frx":1F18
         Stretch         =   -1  'True
         ToolTipText     =   "Top Wall"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   27
         Left            =   150
         Picture         =   "frmCreate.frx":225A
         Stretch         =   -1  'True
         ToolTipText     =   "Left Wall"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   29
         Left            =   450
         Picture         =   "frmCreate.frx":259C
         Stretch         =   -1  'True
         ToolTipText     =   "Right Wall"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   31
         Left            =   450
         Picture         =   "frmCreate.frx":28DE
         Stretch         =   -1  'True
         ToolTipText     =   "Bottom Wall"
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame fraBuildings 
      BackColor       =   &H000080FF&
      Caption         =   "Buildings:"
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
      Height          =   1590
      Left            =   5962
      TabIndex        =   15
      Top             =   30
      Width           =   840
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   18
         Left            =   450
         Picture         =   "frmCreate.frx":2C20
         Stretch         =   -1  'True
         ToolTipText     =   "Bottom Right"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   22
         Left            =   450
         Picture         =   "frmCreate.frx":2F62
         Stretch         =   -1  'True
         ToolTipText     =   "Bottom Wall"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   20
         Left            =   450
         Picture         =   "frmCreate.frx":32A4
         Stretch         =   -1  'True
         ToolTipText     =   "Right Wall"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   19
         Left            =   150
         Picture         =   "frmCreate.frx":35E6
         Stretch         =   -1  'True
         ToolTipText     =   "Left Wall"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   21
         Left            =   150
         Picture         =   "frmCreate.frx":3928
         Stretch         =   -1  'True
         ToolTipText     =   "Top Wall"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   17
         Left            =   150
         Picture         =   "frmCreate.frx":3C6A
         Stretch         =   -1  'True
         ToolTipText     =   "Bottom Left"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   16
         Left            =   450
         Picture         =   "frmCreate.frx":3FAC
         Stretch         =   -1  'True
         ToolTipText     =   "Top Right"
         Top             =   300
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   225
         Index           =   15
         Left            =   150
         Picture         =   "frmCreate.frx":42EE
         Stretch         =   -1  'True
         ToolTipText     =   "Top Left"
         Top             =   300
         Width           =   225
      End
   End
   Begin VB.Frame fraWalkOn 
      BackColor       =   &H000080FF&
      Caption         =   "Walk on:"
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
      Height          =   1290
      Left            =   52
      TabIndex        =   12
      Top             =   60
      Width           =   855
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   7
         Left            =   465
         Picture         =   "frmCreate.frx":4630
         Stretch         =   -1  'True
         ToolTipText     =   "Spikes on Cement"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   5
         Left            =   465
         Picture         =   "frmCreate.frx":4972
         Stretch         =   -1  'True
         ToolTipText     =   "Tile on Cement"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   1
         Left            =   465
         Picture         =   "frmCreate.frx":4CB4
         Stretch         =   -1  'True
         ToolTipText     =   "Cement"
         Top             =   300
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   6
         Left            =   165
         Picture         =   "frmCreate.frx":4FF6
         Stretch         =   -1  'True
         ToolTipText     =   "Spikes on Grass"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   4
         Left            =   165
         Picture         =   "frmCreate.frx":5338
         Stretch         =   -1  'True
         ToolTipText     =   "Tile on Grass"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCommand 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   0
         Left            =   165
         Picture         =   "frmCreate.frx":567A
         Stretch         =   -1  'True
         ToolTipText     =   "Grass"
         Top             =   300
         Width           =   225
      End
   End
   Begin VB.Frame fraMisc 
      BackColor       =   &H000080FF&
      Caption         =   "Misc:"
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
      Height          =   1590
      Left            =   52
      TabIndex        =   13
      Top             =   1470
      Width           =   855
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   23
         Left            =   465
         Picture         =   "frmCreate.frx":59BC
         Stretch         =   -1  'True
         ToolTipText     =   "Locked Door"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   225
         Index           =   8
         Left            =   165
         Picture         =   "frmCreate.frx":5CFE
         Stretch         =   -1  'True
         ToolTipText     =   "Locked Block"
         Top             =   300
         Width           =   225
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   10
         Left            =   165
         Picture         =   "frmCreate.frx":6040
         Stretch         =   -1  'True
         ToolTipText     =   "Bush"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   12
         Left            =   165
         Picture         =   "frmCreate.frx":6390
         Stretch         =   -1  'True
         ToolTipText     =   "Wood"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   14
         Left            =   165
         Picture         =   "frmCreate.frx":66D2
         Stretch         =   -1  'True
         ToolTipText     =   "Wall"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   9
         Left            =   465
         Picture         =   "frmCreate.frx":6A14
         Stretch         =   -1  'True
         ToolTipText     =   "Block"
         Top             =   300
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   11
         Left            =   465
         Picture         =   "frmCreate.frx":6D61
         Stretch         =   -1  'True
         ToolTipText     =   "Brick"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   13
         Left            =   465
         Picture         =   "frmCreate.frx":70A3
         Stretch         =   -1  'True
         ToolTipText     =   "Water"
         Top             =   900
         Width           =   240
      End
   End
   Begin VB.Frame fraGeorge 
      BackColor       =   &H000080FF&
      Caption         =   "Players:"
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
      Height          =   690
      Left            =   52
      TabIndex        =   14
      Top             =   3195
      Width           =   840
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   42
         Left            =   450
         Picture         =   "frmCreate.frx":73E5
         Stretch         =   -1  'True
         ToolTipText     =   "Norman's Starting Position"
         Top             =   300
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   36
         Left            =   120
         Picture         =   "frmCreate.frx":7735
         Stretch         =   -1  'True
         ToolTipText     =   "George's Starting Position"
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Frame fraAdversaries 
      BackColor       =   &H000080FF&
      Caption         =   "Adversaries:"
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   3877
      TabIndex        =   33
      Top             =   4995
      Width           =   1125
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   37
         Left            =   255
         Picture         =   "frmCreate.frx":7A77
         Stretch         =   -1  'True
         ToolTipText     =   "Drone Mouse (Follows wall)"
         Top             =   225
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   38
         Left            =   630
         Picture         =   "frmCreate.frx":7DB9
         Stretch         =   -1  'True
         ToolTipText     =   "Death Mouse (Ultimate AI)"
         Top             =   225
         Width           =   240
      End
   End
   Begin VB.Frame fraItems 
      BackColor       =   &H000080FF&
      Caption         =   "Items:"
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   1882
      TabIndex        =   32
      Top             =   4995
      Width           =   1920
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   41
         Left            =   1590
         Picture         =   "frmCreate.frx":80FB
         Stretch         =   -1  'True
         ToolTipText     =   "Cement Bag(Turns Water into Concrete)"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   40
         Left            =   1290
         Picture         =   "frmCreate.frx":843D
         Stretch         =   -1  'True
         ToolTipText     =   "Bomb (Blows up Blocks)"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   39
         Left            =   990
         Picture         =   "frmCreate.frx":877F
         Stretch         =   -1  'True
         ToolTipText     =   "Clock (Resets Timer to 150)"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   35
         Left            =   690
         Picture         =   "frmCreate.frx":8C05
         Stretch         =   -1  'True
         ToolTipText     =   "Metalic Boots (Walk on Spikes)"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   2
         Left            =   90
         Picture         =   "frmCreate.frx":8F47
         Stretch         =   -1  'True
         ToolTipText     =   "Key on Grass"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgCommand 
         Height          =   240
         Index           =   3
         Left            =   390
         Picture         =   "frmCreate.frx":9289
         Stretch         =   -1  'True
         ToolTipText     =   "Key on Cement"
         Top             =   210
         Width           =   240
      End
   End
   Begin VB.CheckBox chkGrid 
      BackColor       =   &H000080FF&
      Caption         =   "Show Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   52
      TabIndex        =   0
      Top             =   4095
      Value           =   1  'Checked
      Width           =   1050
   End
   Begin VB.Image imgHelp 
      Height          =   330
      Index           =   0
      Left            =   5962
      Picture         =   "frmCreate.frx":95CB
      ToolTipText     =   "Help on Registration"
      Top             =   5265
      Width           =   780
   End
   Begin VB.Label lblBack 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Below George is:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   2737
      TabIndex        =   38
      Top             =   4620
      Width           =   1425
   End
   Begin VB.Label lblCement 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Cement"
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
      Height          =   270
      Left            =   3412
      TabIndex        =   11
      ToolTipText     =   "Ground below you"
      Top             =   4725
      Width           =   765
   End
   Begin VB.Label lblGrass 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grass"
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
      Height          =   270
      Left            =   2662
      TabIndex        =   10
      ToolTipText     =   "Ground below you"
      Top             =   4725
      Width           =   765
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   20
      X1              =   1125
      X2              =   5700
      Y1              =   4635
      Y2              =   4635
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   19
      X1              =   1117
      X2              =   5692
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   18
      X1              =   1117
      X2              =   5692
      Y1              =   4185
      Y2              =   4185
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   17
      X1              =   1117
      X2              =   5692
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   16
      X1              =   1117
      X2              =   5692
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   15
      X1              =   1117
      X2              =   5692
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   14
      X1              =   1117
      X2              =   5692
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   13
      X1              =   1117
      X2              =   5692
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   12
      X1              =   1117
      X2              =   5692
      Y1              =   2835
      Y2              =   2835
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   11
      X1              =   1117
      X2              =   5692
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   10
      X1              =   1117
      X2              =   5692
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   9
      X1              =   1117
      X2              =   5692
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   8
      X1              =   1117
      X2              =   5692
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   7
      X1              =   1117
      X2              =   5692
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   6
      X1              =   1117
      X2              =   5692
      Y1              =   1485
      Y2              =   1485
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   5
      X1              =   1117
      X2              =   5692
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   4
      X1              =   1117
      X2              =   5692
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   1117
      X2              =   5692
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   2
      X1              =   1117
      X2              =   5692
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   1117
      X2              =   5692
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line linColumns 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   1117
      X2              =   5692
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   20
      X1              =   5692
      X2              =   5692
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   19
      X1              =   5467
      X2              =   5467
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   18
      X1              =   5242
      X2              =   5242
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   17
      X1              =   5017
      X2              =   5017
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   16
      X1              =   4792
      X2              =   4792
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   15
      X1              =   4567
      X2              =   4567
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   14
      X1              =   4342
      X2              =   4342
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   13
      X1              =   4117
      X2              =   4117
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   12
      X1              =   3892
      X2              =   3892
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   11
      X1              =   3667
      X2              =   3667
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   10
      X1              =   3442
      X2              =   3442
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   9
      X1              =   3217
      X2              =   3217
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   8
      X1              =   2992
      X2              =   2992
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   7
      X1              =   2767
      X2              =   2767
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   6
      X1              =   2542
      X2              =   2542
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   5
      X1              =   2317
      X2              =   2317
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   4
      X1              =   2092
      X2              =   2092
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   1867
      X2              =   1867
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   2
      X1              =   1642
      X2              =   1642
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   1417
      X2              =   1417
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Line linRows 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   1192
      X2              =   1192
      Y1              =   60
      Y2              =   4635
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   388
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   399
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   398
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   397
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   396
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   395
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   394
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   393
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   392
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   391
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   390
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   389
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   387
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   386
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   385
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   384
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   383
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   382
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   381
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   380
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   379
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   378
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   377
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   376
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   375
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   374
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   373
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   372
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   371
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   370
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   369
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   368
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   367
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   366
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   365
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   364
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   363
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   362
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   361
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   360
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   359
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   358
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   357
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   356
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   355
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   354
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   353
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   352
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   351
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   350
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   349
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   348
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   347
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   346
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   345
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   344
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   343
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   342
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   341
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   340
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   339
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   338
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   337
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   336
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   335
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   334
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   333
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   332
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   331
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   330
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   329
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   328
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   327
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   326
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   325
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   324
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   323
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   322
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   321
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   320
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   319
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   318
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   317
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   316
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   315
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   314
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   313
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   312
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   311
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   310
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   309
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   308
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   307
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   306
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   305
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   304
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   303
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   302
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   301
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   300
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   299
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   298
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   297
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   296
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   295
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   294
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   293
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   292
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   291
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   290
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   289
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   288
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   287
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   286
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   285
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   284
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   283
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   282
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   281
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   280
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   279
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   278
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   277
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   276
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   275
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   274
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   273
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   272
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   271
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   270
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   269
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   268
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   267
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   266
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   265
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   264
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   263
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   262
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   261
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   260
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   259
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   258
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   257
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   256
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   255
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   254
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   253
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   252
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   251
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   250
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   249
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   248
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   247
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   246
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   245
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   244
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   243
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   242
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   241
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   240
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   239
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   238
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   237
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   236
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   235
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   234
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   233
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   232
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   231
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   230
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   229
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   228
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   227
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   226
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   225
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   224
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   223
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   222
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   221
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   220
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   219
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   218
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   217
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   216
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   215
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   214
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   213
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   212
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   211
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   210
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   209
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   208
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   207
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   206
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   205
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   204
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   203
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   202
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   201
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   200
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   199
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   198
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   197
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   196
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   195
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   194
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   193
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   192
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   191
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   190
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   189
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   188
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   187
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   186
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   185
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   184
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   183
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   182
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   181
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   180
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   179
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   178
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   177
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   176
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   175
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   174
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   173
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   172
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   171
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   170
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   169
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   168
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   167
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   166
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   165
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   164
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   163
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   162
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   161
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   160
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   159
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   158
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   157
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   156
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   155
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   154
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   153
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   152
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   151
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   150
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   149
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   148
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   147
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   146
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   145
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   144
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   143
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   142
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   141
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   140
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   139
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   138
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   137
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   136
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   135
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   134
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   133
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   132
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   131
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   130
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   129
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   128
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   127
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   126
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   125
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   124
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   123
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   122
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   121
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   120
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   119
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   118
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   117
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   116
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   115
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   114
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   113
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   112
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   111
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   110
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   109
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   108
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   107
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   106
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   105
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   104
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   103
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   102
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   101
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   100
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   99
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   98
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   97
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   96
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   95
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   94
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   93
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   92
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   91
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   90
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   89
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   88
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   87
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   86
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   85
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   84
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   83
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   82
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   81
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   80
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   79
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   78
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   77
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   76
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   75
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   74
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   73
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   72
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   71
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   70
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   69
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   68
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   67
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   66
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   65
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   64
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   63
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   62
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   61
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   60
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   810
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   59
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   58
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   57
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   56
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   55
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   54
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   53
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   52
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   51
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   50
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   49
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   48
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   47
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   46
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   45
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   44
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   43
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   42
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   41
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   40
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   585
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   39
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   38
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   37
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   36
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   35
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   34
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   33
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   32
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   31
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   30
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   29
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   28
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   27
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   26
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   25
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   24
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   23
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   22
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   21
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   20
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   19
      Left            =   5467
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   18
      Left            =   5242
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   17
      Left            =   5017
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   16
      Left            =   4792
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   15
      Left            =   4567
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   14
      Left            =   4342
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   13
      Left            =   4117
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   12
      Left            =   3892
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   11
      Left            =   3667
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   10
      Left            =   3442
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   9
      Left            =   3217
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   8
      Left            =   2992
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   7
      Left            =   2767
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   6
      Left            =   2542
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   5
      Left            =   2317
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   4
      Left            =   2092
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   3
      Left            =   1867
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   2
      Left            =   1642
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   1
      Left            =   1417
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Image imgMap 
      DragMode        =   1  'Automatic
      Height          =   240
      Index           =   0
      Left            =   1192
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Label lblBackground 
      BackColor       =   &H0080C0FF&
      DragMode        =   1  'Automatic
      Height          =   4515
      Left            =   1192
      TabIndex        =   22
      Top             =   135
      Width           =   4515
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Marque Castle v1.2
'frmCreate
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


'The current picture to create on the map
     Dim strSelectedPic As String
     Dim strUndo(399) As String



'  The Level Number
Private Sub cboLevelNumber_Click()

     'Sets the Best Times File Name
          lblBestTimesDir.Caption = "Lvl" & cboLevelNumber.Text & ".bt"

End Sub

'  Hides/Shows the Grid
Private Sub chkGrid_Click()
  Dim i As Integer
  
    'If there is a Grid
         If chkGrid.Value = 0 Then

               For i = 0 To 20

                    'Hides the Grid
                         linRows(i).Visible = False
                         linColumns(i).Visible = False

               Next

         'There isn't a Grid
               Else

                    For i = 0 To 20

                         'Shows the grid
                              linRows(i).Visible = True
                              linColumns(i).Visible = True

                    Next

         End If

End Sub


'  Whether it is the Last Level or not
Private Sub chkLastLevel_Click()

     'If it isn't the Last Level
          If chkLastLevel.Value = False Then

               'Enable the Next Level Path Input
                    txtNextLevel.Enabled = True

     'If it's the Last Level
          Else

               'Disable the Next Level Path Input
                   txtNextLevel.Enabled = False

          End If

End Sub


Private Sub chkPaintFill_Click()
  Dim i As Integer
  
    For i = 0 To 399
        If chkPaintFill.Value = 0 Then imgMap(i).DragMode = 1
        If chkPaintFill.Value = 1 Then imgMap(i).DragMode = 0
    Next
End Sub

'  Loads a custom scenario
Private Sub cmdLoad_Click()

     'Adjusts the frmFile atributes acordinly
          'Changes Form Caption
               frmFile.Caption = "Data - Load"

          'Changes the Program type Caption
               frmFile.lblTitle.Caption = "Scenario Creation Artist"

         'Changes the Command Button's Caption
               frmFile.cmdSaveLoad.Caption = "Load"

         'Changes the File Type
               frmFile.filFileListBox.Pattern = "*.cus"
               frmFile.txtFileName.Text = "*.cus"
               frmFile.lblFileNameTitle.Caption = "*.cus"

         'Changes the File type Caption
               frmFile.lblFileName = "Scenario to Load:"

         'Centers the File type Caption
               frmFile.lblFileName.Alignment = 2

         'Shows the File Name Input
               frmFile.txtFileName.Visible = False

          'Shows the Form
               frmFile.Show 1

          'Disables the Main Form
'               frmMain.Enabled = False
          'Disables the Scenario Creation Artist Form
'               frmCreate.Enabled = False

End Sub


'  Resumes Creation after Error in Saving
Private Sub cmdResume_Click()

     'Makes a beeping sound
          Beep

     'Hides the Error Frame
          fraError.Visible = False

     'Clears the Error Display List
          cboErrors.Clear
          lblGrass.Enabled = True
          lblCement.Enabled = True
          If fraLevelManager.Width = 6615 Then cmdSave.Enabled = True
          If fraLevelManager.Width = 6615 Then cmdLoad.Enabled = True
          txtLevelTitle.Enabled = True
          If chkLastLevel.Value = 1 Then
              txtNextLevel.Enabled = False
          Else
              txtNextLevel.Enabled = True
          End If
          cboLevelNumber.Enabled = True
          chkLastLevel.Enabled = True
          txtMessage.Enabled = True
          txtAuthor.Enabled = True
          cmdDetails.Enabled = True
          cmdDetails.Caption = "Hide"

End Sub

Private Sub cmdDetails_Click()
     If cmdDetails.Caption = "More" Then
          cmdDetails.Caption = "Hide"
          cmdDetails.Enabled = False
          cmdSave.Enabled = False
          cmdLoad.Enabled = False
          Do While fraLevelManager.Height < 3375
               fraLevelManager.Height = fraLevelManager.Height + 95
               fraLevelManager.Top = fraLevelManager.Top - 95
               DoEvents
          Loop
          chkGrid.Enabled = False
          chkPaintFill.Enabled = False
          cmdUndo.Enabled = False
          txtLevelTitle.Enabled = True
          cboLevelNumber.Enabled = True
          chkLastLevel.Enabled = True
          If chkLastLevel.Value = 1 Then
               txtNextLevel.Enabled = False
          Else
               txtNextLevel.Enabled = True
          End If
          txtMessage.Enabled = True
          txtAuthor.Enabled = True
          cmdDetails.Enabled = True
     Else
          cmdDetails.Caption = "More"
          cmdDetails.Enabled = False
          Do While fraLevelManager.Height > 570
               fraLevelManager.Height = fraLevelManager.Height - 95
               fraLevelManager.Top = fraLevelManager.Top + 95
               DoEvents
          Loop
          chkGrid.Enabled = True
          chkPaintFill.Enabled = True
          cmdUndo.Enabled = True
          txtLevelTitle.Enabled = False
          cboLevelNumber.Enabled = False
          chkLastLevel.Enabled = False
          If chkLastLevel.Value = 1 Then
               txtNextLevel.Enabled = False
          Else
               txtNextLevel.Enabled = True
          End If
          txtMessage.Enabled = False
          txtAuthor.Enabled = False
          cmdDetails.Enabled = True
          If fraLevelManager.Width = 6615 Then cmdSave.Enabled = True
          If fraLevelManager.Width = 6615 Then cmdLoad.Enabled = True
     End If
End Sub

Private Sub cmdUndo_Click()
  Dim i As Integer

  On Error Resume Next

     Dim strTempUndo(399) As String
    ' Store what you have for another undo
    For i = 0 To 399
        strTempUndo(i) = strCell(i)
    Next
    ' Restore original
    For i = 0 To 399
        strCell(i) = strUndo(i)
        imgMap(i).Picture = frmMain.pic(CInt(strUndo(i))).Picture
    Next
    ' Set new undo
    For i = 0 To 399
        strUndo(i) = strTempUndo(i)
    Next
End Sub

'  Entry to the Level Creation Artist
Private Sub Form_Load()
  Dim i As Integer
 
      'Sets the defaults
          strSelectedPic = "000"
          imgCommand(0).BorderStyle = 1
          If blnSecret = False Then
              For i = 0 To 399
                  imgMap(i).Enabled = True
              Next
          End If
          ' All grass map
         For i = 0 To 399
            strCell(i) = "000"
            imgMap(i).Picture = frmMain.pic(0).Picture
         Next

End Sub


'Exiting the Level Creation Artist
Private Sub Form_Unload(Cancel As Integer)

    frmMain.Enabled = True

End Sub

'Changing the picture you draw with to what you've clicked on
Private Sub imgCommand_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  
     'If the Level Manager isn't maximized then
          If cmdDetails.Caption = "More" And fraError.Visible = False Then

         'Unselects all
               For i = 0 To 42
                    imgCommand(i).BorderStyle = 0
               Next

         'Selects the one you've clicked
               imgCommand(Index).BorderStyle = 1

         'Assigns strSelectedPic with the Picture you've selcted
          strSelectedPic = imgCommand(Index).Index
          strSelectedPic = "00" & strSelectedPic
          If Len(strSelectedPic) > 3 Then
              strSelectedPic = Mid$(strSelectedPic, 2, 3)
          End If
          If strSelectedPic = "036" Then
              strSelectedPic = "093"
          ElseIf strSelectedPic = "037" Then
              strSelectedPic = "991"
          ElseIf strSelectedPic = "038" Then
              strSelectedPic = "992"
          ElseIf strSelectedPic = "039" Then
              strSelectedPic = "036"
          ElseIf strSelectedPic = "040" Then
              strSelectedPic = "037"
          ElseIf strSelectedPic = "041" Then
               strSelectedPic = "038"
          ElseIf strSelectedPic = "042" Then
               strSelectedPic = "893"
          End If
        
     End If

End Sub


Private Sub imgHelp_Click(Index As Integer)
    If frmMain.mnuSound.Checked = True Then PlaySound 0, App.Path & "\Beep.wav"
    If Index = 0 Then
        MsgBox "HELP!" & vbNewLine & vbNewLine & "The Level Creation Artist alows you to create your own Marque Castle game." & vbNewLine & "  >> Select one of the pictures on the left or right" & vbNewLine & "  >> Select where you want to place it on your map, by clicking that map area" & vbNewLine & "  >> Be sure you fill in all required info. for your level, displayed under ''More''", vbInformation, "Marque Castle"
    Else
        MsgBox "HELP!" & vbNewLine & vbNewLine & "There are ''Rules'' as to what you can and can't save." & vbNewLine & vbNewLine & "  >> Every square must be filled" & vbNewLine & "  >> Keys: 1 or more" & vbNewLine & "  >> Doors: 1" & vbNewLine & "  >> George: No more than 1" & vbNewLine & "  >> Norman: 1" & vbNewLine & "  >> Toggle Blocks: No more than 250 (On or Off)", vbInformation, "Marque Castle"
    End If
End Sub

Private Sub imgMap_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
     'If the Level Manager is maximized then
     If cmdDetails.Caption = "Hide" Then Exit Sub

    ' Using regular draw
    If chkPaintFill.Value = 0 Then
        ' First, store the undo
        StoreUndo
        ' Updates the strCell information about the objects
        strCell(imgMap(Index).Index) = strSelectedPic
        ' Places the selected picture in the selected block
        imgMap(Index).Picture = frmMain.pic(CInt(strSelectedPic)).Picture
    End If
End Sub

' The PaintFill subroutine
Private Sub PaintFill(Coords As Integer, ChangeWhat As String)
    ' Possible stack space error
    On Error GoTo Hell
    ' 1st, change the starting coords
    strCell(Coords) = strSelectedPic
    imgMap(Coords).Picture = frmMain.pic(CInt(strSelectedPic)).Picture
    
    ' Then check the one above it
    If (Coords - 20) > 0 Then
        ' Not on an edge
        If strCell(Coords - 20) = ChangeWhat Then If Coords - 20 > 19 Then PaintFill Coords - 20, ChangeWhat
    End If
    
    ' ... below it
    If (Coords + 20) < 399 Then
        ' Not on an edge
        If strCell(Coords + 20) = ChangeWhat Then If Coords + 20 < 380 Then PaintFill Coords + 20, ChangeWhat
    End If
    
    ' .. left of it
    If Coords - 1 >= 0 Then
        ' Not on an edge
        If strCell(Coords - 1) = ChangeWhat Then If Coords - 1 Mod 20 <> 0 Then PaintFill Coords - 1, ChangeWhat
    End If
    
    ' ... right of it
    If Coords + 1 <= 399 Then
        ' Not on an edge
        If strCell(Coords + 1) = ChangeWhat Then If Coords + 1 Mod 20 <> 19 Then PaintFill Coords + 1, ChangeWhat
    End If
Hell:
    ' Do nothing
End Sub

Private Sub imgMap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If the Level Manager is maximized then
    If cmdDetails.Caption = "Hide" Then Exit Sub
    ' Left clicked, and you're using the paint fill
    If Button = 1 And chkPaintFill.Value = 1 And strCell(imgMap(Index).Index) <> strSelectedPic Then
        ' First, store the undo
        StoreUndo
        Dim strOriginal As String
        strOriginal = strCell(imgMap(Index).Index)
        PaintFill Index, strOriginal
    End If
End Sub


'Making the default pictures Grass
Private Sub lblGrass_Click()

    'Toggles between Grass and Cement as the primary object
    lblGrass.BorderStyle = 1: lblCement.BorderStyle = 0

End Sub


'Making the default pictures Cement
Private Sub lblCement_Click()

    'Toggles between Grass and Cement as the primary object
    lblGrass.BorderStyle = 0: lblCement.BorderStyle = 1

End Sub


'Saving your current project
Private Sub cmdSave_Click()

    'Adjusts the frmFile atributes acordinly
        frmFile.Caption = "Data - Save"
        frmFile.lblTitle.Caption = "Scenario Creation Artist"
        frmFile.cmdSaveLoad.Caption = "Save"
        frmFile.filFileListBox.Pattern = "*.cus"
        frmFile.txtFileName.Text = "*.cus"
        frmFile.lblFileName = "Scenario to Save:"
        frmFile.lblFileName.Alignment = 1
        frmFile.Show 1
'        frmMain.Enabled = False
'        frmCreate.Enabled = False

End Sub


Private Sub StoreUndo()
  Dim i As Integer

    ' Store the original map
    For i = 0 To 399
        strUndo(i) = strCell(i)
    Next
    ' Enable the undo button
    cmdUndo.Enabled = True
End Sub
