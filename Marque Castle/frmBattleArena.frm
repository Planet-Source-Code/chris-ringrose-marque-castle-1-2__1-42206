VERSION 5.00
Begin VB.Form frmBattleArena 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13905
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7680
      Left            =   1552
      TabIndex        =   1
      Top             =   810
      Width           =   10800
      Begin VB.Frame fraPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Player One (George):"
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1875
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   1170
         Width           =   2115
         Begin VB.Label lblWins 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   315
            Index           =   0
            Left            =   945
            TabIndex        =   18
            Top             =   1410
            Width           =   345
         End
         Begin VB.Label lblDividor 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   " /"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   315
            Index           =   0
            Left            =   1290
            TabIndex        =   17
            Top             =   1410
            Width           =   225
         End
         Begin VB.Label lblWinsMax 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   315
            Index           =   0
            Left            =   1515
            TabIndex        =   16
            Top             =   1410
            Width           =   345
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   1
            X1              =   1845
            X2              =   1860
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   0
            X1              =   945
            X2              =   960
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Image imgCement 
            Appearance      =   0  'Flat
            Height          =   435
            Index           =   0
            Left            =   150
            Stretch         =   -1  'True
            Top             =   240
            Width           =   435
         End
         Begin VB.Label lblX 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   225
            Index           =   0
            Left            =   615
            TabIndex        =   15
            Top             =   420
            Width           =   225
         End
         Begin VB.Label lblCementBags 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   315
            Index           =   0
            Left            =   825
            TabIndex        =   14
            ToolTipText     =   "Number of Bags of Cement"
            Top             =   360
            Width           =   315
         End
         Begin VB.Image imgBoots 
            Appearance      =   0  'Flat
            Height          =   435
            Index           =   0
            Left            =   420
            Stretch         =   -1  'True
            Top             =   735
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Image imgBomb 
            Appearance      =   0  'Flat
            Height          =   435
            Index           =   0
            Left            =   1110
            Stretch         =   -1  'True
            Top             =   735
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Line linDividor01 
            BorderColor     =   &H000080FF&
            Index           =   0
            X1              =   150
            X2              =   1950
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label lblWinsLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Wins:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   255
            Index           =   0
            Left            =   345
            TabIndex        =   13
            Top             =   1485
            Width           =   615
         End
         Begin VB.Line linWins 
            BorderColor     =   &H000080FF&
            Index           =   0
            X1              =   960
            X2              =   1845
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   4
            X1              =   15
            X2              =   30
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   5
            X1              =   2085
            X2              =   2100
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   6
            X1              =   15
            X2              =   30
            Y1              =   1845
            Y2              =   1845
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   7
            X1              =   2085
            X2              =   2100
            Y1              =   1845
            Y2              =   1845
         End
      End
      Begin VB.Frame fraBattleArena 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6300
         Left            =   2175
         TabIndex        =   9
         Top             =   1335
         Width           =   6300
         Begin VB.Frame fraCountDown 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1590
            Left            =   1755
            TabIndex        =   10
            Top             =   1620
            Width           =   3195
            Begin VB.Timer tmrCountDown 
               Enabled         =   0   'False
               Interval        =   900
               Left            =   3015
               Top             =   1350
            End
            Begin VB.Label lblCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H000080FF&
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   72
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   1470
               Left            =   0
               TabIndex        =   11
               Top             =   60
               Width           =   3375
            End
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   315
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   630
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   945
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   8
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   9
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   10
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   11
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   12
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   13
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   14
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   15
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   16
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   17
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   18
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   19
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   20
            Left            =   0
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   21
            Left            =   315
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   22
            Left            =   630
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   23
            Left            =   945
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   24
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   25
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   26
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   27
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   28
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   29
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   30
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   31
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   32
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   33
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   34
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   35
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   36
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   37
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   38
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   39
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   315
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   40
            Left            =   0
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   41
            Left            =   315
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   42
            Left            =   630
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   43
            Left            =   945
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   44
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   45
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   46
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   47
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   48
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   49
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   50
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   51
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   52
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   53
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   54
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   55
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   56
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   57
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   58
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   59
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   630
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   60
            Left            =   0
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   61
            Left            =   315
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   62
            Left            =   630
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   63
            Left            =   945
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   64
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   65
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   66
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   67
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   68
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   69
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   70
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   71
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   72
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   73
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   74
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   75
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   76
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   77
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   78
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   79
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   945
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   80
            Left            =   0
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   81
            Left            =   315
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   82
            Left            =   630
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   83
            Left            =   945
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   84
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   85
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   86
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   87
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   88
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   89
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   90
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   91
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   92
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   93
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   94
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   95
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   96
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   97
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   98
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   99
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   100
            Left            =   0
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   101
            Left            =   315
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   102
            Left            =   630
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   103
            Left            =   945
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   104
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   105
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   106
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   107
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   108
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   109
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   110
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   111
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   112
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   113
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   114
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   115
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   116
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   117
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   118
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   119
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   1575
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   120
            Left            =   0
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   121
            Left            =   315
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   122
            Left            =   630
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   123
            Left            =   945
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   124
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   125
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   126
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   127
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   128
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   129
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   130
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   131
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   132
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   133
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   134
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   135
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   136
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   137
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   138
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   139
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   1890
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   140
            Left            =   0
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   141
            Left            =   315
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   142
            Left            =   630
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   143
            Left            =   945
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   144
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   145
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   146
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   147
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   148
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   149
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   150
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   151
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   152
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   153
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   154
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   155
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   156
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   157
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   158
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   159
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   2205
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   160
            Left            =   0
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   161
            Left            =   315
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   162
            Left            =   630
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   163
            Left            =   945
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   164
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   165
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   166
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   167
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   168
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   169
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   170
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   171
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   172
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   173
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   174
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   175
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   176
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   177
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   178
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   179
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   180
            Left            =   0
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   181
            Left            =   315
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   182
            Left            =   630
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   183
            Left            =   945
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   184
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   185
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   186
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   187
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   188
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   189
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   190
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   191
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   192
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   193
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   194
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   195
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   196
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   197
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   198
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   199
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   2835
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   200
            Left            =   0
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   201
            Left            =   315
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   202
            Left            =   630
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   203
            Left            =   945
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   204
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   205
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   206
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   207
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   208
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   209
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   210
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   211
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   212
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   213
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   214
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   215
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   216
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   217
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   218
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   219
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   3150
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   220
            Left            =   0
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   221
            Left            =   315
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   222
            Left            =   630
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   223
            Left            =   945
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   224
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   225
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   226
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   227
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   228
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   229
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   230
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   231
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   232
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   233
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   234
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   235
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   236
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   237
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   238
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   239
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   3465
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   240
            Left            =   0
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   241
            Left            =   315
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   242
            Left            =   630
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   243
            Left            =   945
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   244
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   245
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   246
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   247
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   248
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   249
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   250
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   251
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   252
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   253
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   254
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   255
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   256
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   257
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   258
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   259
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   3780
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   260
            Left            =   0
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   261
            Left            =   315
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   262
            Left            =   630
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   263
            Left            =   945
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   264
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   265
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   266
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   267
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   268
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   269
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   270
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   271
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   272
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   273
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   274
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   275
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   276
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   277
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   278
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   279
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   4095
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   280
            Left            =   0
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   281
            Left            =   315
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   282
            Left            =   630
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   283
            Left            =   945
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   284
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   285
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   286
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   287
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   288
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   289
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   290
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   291
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   292
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   293
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   294
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   295
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   296
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   297
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   298
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   299
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   4410
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   300
            Left            =   0
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   301
            Left            =   315
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   302
            Left            =   630
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   303
            Left            =   945
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   304
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   305
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   306
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   307
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   308
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   309
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   310
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   311
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   312
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   313
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   314
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   315
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   316
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   317
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   318
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   319
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   4725
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   320
            Left            =   0
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   321
            Left            =   315
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   322
            Left            =   630
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   323
            Left            =   945
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   324
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   325
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   326
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   327
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   328
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   329
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   330
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   331
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   332
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   333
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   334
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   335
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   336
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   337
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   338
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   339
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   340
            Left            =   0
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   341
            Left            =   315
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   342
            Left            =   630
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   343
            Left            =   945
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   344
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   345
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   346
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   347
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   348
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   349
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   350
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   351
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   352
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   353
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   354
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   355
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   356
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   357
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   358
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   359
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   5355
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   360
            Left            =   0
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   361
            Left            =   315
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   362
            Left            =   630
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   363
            Left            =   945
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   364
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   365
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   366
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   367
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   368
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   369
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   370
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   371
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   372
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   373
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   374
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   375
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   376
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   377
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   378
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   379
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   5670
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   380
            Left            =   0
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   381
            Left            =   315
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   382
            Left            =   630
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   383
            Left            =   945
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   384
            Left            =   1260
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   385
            Left            =   1575
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   386
            Left            =   1890
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   387
            Left            =   2205
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   388
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   389
            Left            =   2835
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   390
            Left            =   3150
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   391
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   392
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   393
            Left            =   4095
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   394
            Left            =   4410
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   395
            Left            =   4725
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   396
            Left            =   5040
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   397
            Left            =   5355
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   398
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBtlMap 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   399
            Left            =   5985
            Stretch         =   -1  'True
            Top             =   5985
            Width           =   315
         End
         Begin VB.Image imgBoom 
            Appearance      =   0  'Flat
            Height          =   915
            Left            =   30
            Stretch         =   -1  'True
            Top             =   30
            Visible         =   0   'False
            Width           =   885
         End
      End
      Begin VB.Frame fraPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Player Two (Norman):"
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1875
         Index           =   1
         Left            =   8700
         TabIndex        =   2
         Top             =   1170
         Width           =   2115
         Begin VB.Line linWins 
            BorderColor     =   &H000080FF&
            Index           =   1
            X1              =   960
            X2              =   1845
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Label lblWinsLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Wins:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   255
            Index           =   1
            Left            =   345
            TabIndex        =   8
            Top             =   1485
            Width           =   615
         End
         Begin VB.Label lblWinsMax 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   315
            Index           =   1
            Left            =   1515
            TabIndex        =   7
            Top             =   1410
            Width           =   345
         End
         Begin VB.Label lblDividor 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   " /"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   315
            Index           =   1
            Left            =   1290
            TabIndex        =   6
            Top             =   1410
            Width           =   225
         End
         Begin VB.Label lblWins 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   315
            Index           =   1
            Left            =   945
            TabIndex        =   5
            Top             =   1410
            Width           =   345
         End
         Begin VB.Line linDividor01 
            BorderColor     =   &H000080FF&
            Index           =   1
            X1              =   150
            X2              =   1950
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Image imgBomb 
            Appearance      =   0  'Flat
            Height          =   435
            Index           =   1
            Left            =   1110
            Stretch         =   -1  'True
            Top             =   735
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Image imgBoots 
            Appearance      =   0  'Flat
            Height          =   435
            Index           =   1
            Left            =   420
            Stretch         =   -1  'True
            Top             =   735
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label lblCementBags 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   315
            Index           =   1
            Left            =   825
            TabIndex        =   4
            ToolTipText     =   "Number of Bags of Cement"
            Top             =   360
            Width           =   315
         End
         Begin VB.Label lblX 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   225
            Index           =   1
            Left            =   615
            TabIndex        =   3
            Top             =   420
            Width           =   225
         End
         Begin VB.Image imgCement 
            Appearance      =   0  'Flat
            Height          =   435
            Index           =   1
            Left            =   150
            Stretch         =   -1  'True
            Top             =   240
            Width           =   435
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   2
            X1              =   945
            X2              =   960
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   3
            X1              =   1845
            X2              =   1860
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   8
            X1              =   15
            X2              =   30
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   9
            X1              =   2085
            X2              =   2100
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   10
            X1              =   15
            X2              =   30
            Y1              =   1845
            Y2              =   1845
         End
         Begin VB.Line linDot 
            BorderColor     =   &H00000000&
            Index           =   11
            X1              =   2085
            X2              =   2100
            Y1              =   1845
            Y2              =   1845
         End
      End
      Begin VB.Image imgTitle 
         Height          =   1110
         Left            =   1770
         Picture         =   "frmBattleArena.frx":0000
         Top             =   0
         Width           =   7275
      End
   End
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   765
   End
End
Attribute VB_Name = "frmBattleArena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Marque Castle v1.2
'frmBattleArena
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



'George's Position
     Dim intGeorgePos As Integer

'Norman's Position
     Dim intNormanPos As Integer

'George's Ground
     Dim intGeorgeGround As Integer

'Norman's Ground
     Dim intNormanGround As Integer

'George's Number of Bags of Cement
     Dim intGeorgesCement As Integer

'Norman's Number of Bags of Cement
     Dim intNormansCement As Integer

'Player's wins (0-George; 1-Norman)
     Dim intWinCount(1) As Integer

'The Level Data
     Dim intBattleData(399) As Integer

'The Default Ground
     Dim intDefGround As Integer


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

          If fraCountDown.Visible = True Then Exit Sub

     'What Key was pressed
          Select Case KeyCode

               'If you've pressed <UP>
                    Case Is = vbKeyUp

                         'Checks World Border
                              If intGeorgePos > 19 Then

                                   'Moves George to the appropriate possition
                                        GeorgeM -20, 91, True, intGeorgePos

                              End If

               'If you've pressed <W>
                    Case Is = vbKeyW

                         'Checks World Border
                              If intNormanPos > 19 Then

                                   'Moves George to the appropriate possition
                                        GeorgeM -20, 891, False, intNormanPos

                              End If

               'If you've pressed <DOWN>
                    Case Is = vbKeyDown

                         'Checks World Border
                              If intGeorgePos < 380 Then

                                   'Moves George to the appropriate possition
                                        GeorgeM 20, 93, True, intGeorgePos

                              End If

               'If you've pressed <S>
                    Case Is = vbKeyS

                         'Checks World Border
                              If intNormanPos < 380 Then

                                   'Moves George to the appropriate possition
                                        GeorgeM 20, 893, False, intNormanPos

                              End If

               'If you've pressed <RIGHT>
                    Case Is = vbKeyRight

                         'Checks World Border
                              If intGeorgePos Mod 20 <> 19 Then

                                   'Moves George to the appropriate possition
                                        GeorgeM 1, 92, True, intGeorgePos

                              End If

               'If you've pressed <D>
                    Case Is = vbKeyD

                         'Checks World Border
                              If intNormanPos Mod 20 <> 19 Then

                                   'Moves George to the appropriate possition
                                        GeorgeM 1, 892, False, intNormanPos

                              End If

               'If you've pressed <LEFT>
                    Case Is = vbKeyLeft

                         'Checks World Border
                              If intGeorgePos Mod 20 <> 0 Then

                                   'Moves George to the appropriate possition
                                      GeorgeM -1, 94, True, intGeorgePos

                              End If

               'If you've pressed <A>
                    Case Is = vbKeyA

                         'Checks World Border
                              If intNormanPos Mod 20 <> 0 Then

                                   'Moves George to the appropriate possition
                                      GeorgeM -1, 894, False, intNormanPos

                              End If

              'If you've pressed <Space Bar>
                      Case Is = vbKeySpace

                          'Checks to see if you have a Bomb
                              If imgBomb(0).Visible = True Then

                                   'Removes the Bomb pic
                                        imgBomb(0).Visible = False

                                   If frmMain.mnuSound.Checked = True Then
                                        'Makes an explosion sound
                                             PlaySound 0, App.Path & "\UseBomb.wav"
                                   End If

                                  'Makes the Explosion
                                        'Explosion intGeorgePos

                              End If

                    End Select

End Sub


'  Code on Startup
Private Sub Form_Load()
  Dim z As String
  Dim p As Integer
  Dim i As Integer

  lblWinsMax(0).Caption = CInt(Mid(frmMain.cboStyle.Text, 1, 2))
  lblWinsMax(1).Caption = CInt(Mid(frmMain.cboStyle.Text, 1, 2))


    fraMain.Left = (frmBattleArena.Width - frmBattleArena.fraMain.Left) \ 2
    fraMain.Top = ((frmBattleArena.Height - frmBattleArena.fraMain.Top) \ 2) \ 2

     'Starts to Play the Theme Song
          If frmMain.mnuSound.Checked = True Then frmSplash.medMidi.URL = App.Path & "\Marque Theme.mid"

     'Centres the Playing Area
        fraMain.Left = (frmBattleArena.ScaleWidth - fraMain.Width) / 2
        fraMain.Top = (frmBattleArena.ScaleHeight - fraMain.Height) / 2

     'Loads the Info into the Levels
          'Sets the Value of fraPictures to store whether Reading or Writing
               frmMain.fraPictures.Caption = "Read"
          'Checks for Errors while Opening
               FileExists App.Path & "\Battle.btl"

     'Stores the Level Data
          Input #1, z, intGeorgeGround

     'Sets the Value of Norman's Ground
          intNormanGround = intGeorgeGround

     'Loads the Pictures into their place on the map
                p = 1
                For i = 0 To 399
                    imgBtlMap(i).Picture = frmMain.pic(CInt(Mid(z, p, 3))).Picture
                    intBattleData(i) = CInt(Mid(z, p, 3))
                    If intBattleData(i) = 93 Then intGeorgePos = i
                    If intBattleData(i) = 893 Then intNormanPos = i
                    p = p + 3
                Next

          For i = 0 To 1
               imgCement(i).Picture = frmMain.pic(38).Picture
               imgBoots(i).Picture = frmMain.pic(35).Picture
               imgBomb(i).Picture = frmMain.pic(37).Picture
          Next
          imgBoom.Picture = LoadPicture(strSkinDir & "\Explosion.gif")

     'Enables the CountDown Timer
          tmrCountDown.Enabled = True

     'Sets the Default Ground
          intDefGround = intGeorgeGround

End Sub

Private Sub lblQuit_Click()

     If MsgBox("Are you sure you would like to leave the Marque Battle Arena?", vbYesNo, "Marque Castle") = vbYes Then
          'Stops the Theme
               frmSplash.medMidi.URL = ""
          'Unloads the Battle Arena
               Unload frmBattleArena
     End If

End Sub


Private Sub tmrCountDown_Timer()

     If frmMain.mnuSound.Checked = True Then
          'Makes an beeping sound
               PlaySound 0, App.Path & "\Beep.wav"
     End If

     If lblCount.Caption = "5" Then
          lblCount.Caption = "4"
     ElseIf lblCount.Caption = "4" Then
          lblCount.Caption = "3"
     ElseIf lblCount.Caption = "3" Then
          lblCount.Caption = "2"
     ElseIf lblCount.Caption = "2" Then
          lblCount.Caption = "1"
     ElseIf lblCount.Caption = "1" Then
          fraCountDown.Visible = False
          tmrCountDown.Enabled = False
     End If

End Sub


Private Sub BackGeorge(MoveTo As Integer, PlayerPic As Integer)

     'Removes George
          'Back to the Below Picture (Removing George)
               imgBtlMap(intGeorgePos).Picture = frmMain.pic(intGeorgeGround).Picture
               intBattleData(intGeorgePos) = intGeorgeGround

     'Places the new Possition and Picture (Adding George)
          intGeorgePos = intGeorgePos + MoveTo
          intGeorgeGround = intBattleData(intGeorgePos)
          intBattleData(intGeorgePos) = PlayerPic

          imgBtlMap(intGeorgePos).Picture = frmMain.pic(PlayerPic).Picture

End Sub


Private Sub BackNorman(MoveTo As Integer, PlayerPic As Integer)

     'Removes Norman
          'Back to the Below Picture (Removing Norman)
               imgBtlMap(intNormanPos).Picture = frmMain.pic(intNormanGround).Picture
               intBattleData(intNormanPos) = intNormanGround

     'Places the new Possition and Picture (Adding Norman)
          intNormanPos = intNormanPos + MoveTo
          intNormanGround = intBattleData(intNormanPos)
          intBattleData(intNormanPos) = PlayerPic

          imgBtlMap(intNormanPos).Picture = frmMain.pic(PlayerPic).Picture

End Sub


Private Sub GeorgeM(MoveTo As Integer, PlayerPic As Integer, George As Boolean, intpositions)
  Dim p As Integer
  Dim i As Integer
  Dim r As Integer

     'Either <Grass>, <Cement> or <Toggle Block OFF>
          If intBattleData(intpositions + MoveTo) < 2 Or intBattleData(intpositions + MoveTo) = 33 Then

               'Moves
                    If George = True Then
                      BackGeorge MoveTo, PlayerPic
                    Else
                      BackNorman MoveTo, PlayerPic
                    End If

     '<Tile on Grass> or <Tile on Cement>
          ElseIf intBattleData(intpositions + MoveTo) = 4 Or intBattleData(intpositions + MoveTo) = 5 Then

               If frmMain.mnuSound.Checked = True Then
                    'Makes a stepped on tile sound
                         PlaySound 0, App.Path & "\Tile.wav"
               End If

               'Moves
                    If George = True Then
                      BackGeorge MoveTo, PlayerPic
                    Else
                      BackNorman MoveTo, PlayerPic
                    End If

               'If Norman is on a <<OFF>> Toggle Block
                    If intNormanGround = 33 Then

                         'Shows Dead Norman Pic
                              imgBtlMap(intNormanPos).Picture = frmMain.pic(895).Picture

                         'Victory for George

               'If George is on a <<OFF>> Toggle Block
                    ElseIf intGeorgeGround = 33 Then

                         'Shows Dead George Pic
                              imgBtlMap(intGeorgePos).Picture = frmMain.pic(95).Picture
                    
                    End If
     
                    p = 1
                    For i = 0 To 399

                         'Found a Toggle Block <<ON>>
                              If intBattleData(i) = 34 Then

                                   'Reverses the Blocks (If On then Off)
                                        intBattleData(i) = 33
                                        imgBtlMap(i).Picture = frmMain.pic(33).Picture

                         'Found a Toggle Block <<OFF>>
                              ElseIf intBattleData(i) = 33 Then

                                   'Reverses the Blocks (If Off then On)
                                        intBattleData(i) = 34
                                        imgBtlMap(i).Picture = frmMain.pic(34).Picture

                              End If
               Next

     'Either <Spikes on Grass> or <Spikes on Cement>
          ElseIf intBattleData(intpositions + MoveTo) = 6 Or intBattleData(intpositions + MoveTo) = 7 Then

               'Moves
                    If George = True Then
                         BackGeorge MoveTo, PlayerPic
                    Else
                         BackNorman MoveTo, PlayerPic
                    End If

               If frmMain.mnuSound.Checked = True Then
                    'Makes a stepped on spikes sound
                         r = RndBetween(1, 2)
                         If r < 1.5 Then
                              PlaySound 0, App.Path & "\Spikes1.wav"
                         Else
                              PlaySound 0, App.Path & "\Spikes2.wav"
                         End If
               End If

                'If you don't have the metallic Boots then
                    If imgBoots(0).Visible = False Then

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a defeat sound
                                   PlaySound 0, App.Path & "\Defeat.wav"
                         End If

                              'Lost 0

                    End If

     '<Water>
          ElseIf intBattleData(intpositions + MoveTo) = 13 Then

               'Moves
                    If George = True Then
                         'Removes George
                              'Back to the Below Picture (Removing George)
                                   imgBtlMap(intpositions).Picture = frmMain.pic(intGeorgeGround).Picture
                                   intBattleData(intpositions) = intGeorgeGround

                         'Places the new Possition and Picture (Adding George)
                              intpositions = intpositions + MoveTo
                              intGeorgeGround = 1
                              intBattleData(intpositions) = PlayerPic

                              imgBtlMap(intpositions).Picture = frmMain.pic(PlayerPic).Picture
                    Else
                         'Removes Norman
                              'Back to the Below Picture (Removing Norman)
                                   imgBtlMap(intNormanPos).Picture = frmMain.pic(intNormanGround).Picture
                                   intBattleData(intNormanPos) = intNormanGround
          
                         'Places the new Possition and Picture (Adding George)
                              intNormanPos = intNormanPos + MoveTo
                              intNormanGround = 1
                              intBattleData(intNormanPos) = PlayerPic
          
                              imgBtlMap(intNormanPos).Picture = frmMain.pic(PlayerPic).Picture

                    End If

               'If you have a Cement Bag
                    If intGeorgesCement > 0 Then

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a stepped on water sound
                                   PlaySound 0, App.Path & "\Water.wav"
                         End If

                              'Changes the Variable (One less Cement Bag)
                                   intGeorgesCement = intGeorgesCement - 1
                                   lblCementBags(0).Caption = intGeorgesCement

                         'If you dont have a Cement Bag
                              Else

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a defeat sound
                                   PlaySound 0, App.Path & "\Defeat.wav"
                         End If

                                   Defeat

                    End If

      '<Metallic Boots>
          ElseIf intBattleData(intpositions + MoveTo) = 35 Then

               'Ensures that you don't already have the Boots
                    If imgBoots(0).Visible = False Then

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a picking up boots sound
                                   PlaySound 0, App.Path & "\GotBoots.wav"
                         End If

                         'Shows the Boots Picture in the Items menu
                               imgBoots(0).Visible = True
          
                         'Moves
                              If George = True Then
                                   'Removes George
                                        'Back to the Below Picture (Removing George)
                                             imgBtlMap(intpositions).Picture = frmMain.pic(intGeorgeGround).Picture
                                             intBattleData(intpositions) = intGeorgeGround

                                   'Places the new Possition and Picture (Adding George)
                                        intpositions = intpositions + MoveTo
                                        intGeorgeGround = intDefGround
                                        intBattleData(intpositions) = PlayerPic

                                        imgBtlMap(intpositions).Picture = frmMain.pic(PlayerPic).Picture
                              Else
                                   'Removes Norman
                                        'Back to the Below Picture (Removing Norman)
                                             imgBtlMap(intNormanPos).Picture = frmMain.pic(intNormanGround).Picture
                                             intBattleData(intNormanPos) = intNormanGround

                                   'Places the new Possition and Picture (Adding George)
                                        intNormanPos = intNormanPos + MoveTo
                                        intNormanGround = intDefGround
                                        intBattleData(intNormanPos) = PlayerPic
     
                                        imgBtlMap(intNormanPos).Picture = frmMain.pic(PlayerPic).Picture

                              End If

                    End If

      '<Bomb>
            ElseIf intBattleData(intpositions + MoveTo) = 37 Then

               'Makes sure that you don't already have a Bomb
                    If imgBomb(0).Visible = False Then

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a picking up bomb sound
                                   PlaySound 0, App.Path & "\GotBomb.wav"
                         End If

                         'Shows the Bomb Picture in Items menu
                              imgBomb(0).Visible = True

                         'Moves
                              If George = True Then
                                   'Removes George
                                        'Back to the Below Picture (Removing George)
                                             imgBtlMap(intpositions).Picture = frmMain.pic(intGeorgeGround).Picture
                                             intBattleData(intpositions) = intGeorgeGround

                                   'Places the new Possition and Picture (Adding George)
                                        intpositions = intpositions + MoveTo
                                        intGeorgeGround = intDefGround
                                        intBattleData(intpositions) = PlayerPic

                                        imgBtlMap(intpositions).Picture = frmMain.pic(PlayerPic).Picture
                              Else
                                   'Removes Norman
                                        'Back to the Below Picture (Removing Norman)
                                             imgBtlMap(intNormanPos).Picture = frmMain.pic(intNormanGround).Picture
                                             intBattleData(intNormanPos) = intNormanGround

                                   'Places the new Possition and Picture (Adding George)
                                        intNormanPos = intNormanPos + MoveTo
                                        intNormanGround = intDefGround
                                        intBattleData(intNormanPos) = PlayerPic

                                        imgBtlMap(intNormanPos).Picture = frmMain.pic(PlayerPic).Picture

                              End If
               
               End If


     '<Cement Bag>
          ElseIf intBattleData(intpositions + MoveTo) = 38 Then

               If frmMain.mnuSound.Checked = True Then
                    'Makes a picking up cement bag sound
                         PlaySound 0, App.Path & "\GotClock.wav"
               End If

               'Changes the Variable (You have one more Cement Bag)
                    intGeorgesCement = intGeorgesCement + 1
                    lblCementBags(0).Caption = intGeorgesCement

                         'Moves
                              If George = True Then
                                   'Removes George
                                        'Back to the Below Picture (Removing George)
                                             imgBtlMap(intpositions).Picture = frmMain.pic(intGeorgeGround).Picture
                                             intBattleData(intpositions) = intGeorgeGround

                                   'Places the new Possition and Picture (Adding George)
                                        intpositions = intpositions + MoveTo
                                        intGeorgeGround = intDefGround
                                        intBattleData(intpositions) = PlayerPic

                                        imgBtlMap(intpositions).Picture = frmMain.pic(PlayerPic).Picture
                              Else
                                   'Removes Norman
                                        'Back to the Below Picture (Removing Norman)
                                             imgBtlMap(intNormanPos).Picture = frmMain.pic(intNormanGround).Picture
                                             intBattleData(intNormanPos) = intNormanGround

                                   'Places the new Possition and Picture (Adding George)
                                        intNormanPos = intNormanPos + MoveTo
                                        intNormanGround = intDefGround
                                        intBattleData(intNormanPos) = PlayerPic

                                        imgBtlMap(intNormanPos).Picture = frmMain.pic(PlayerPic).Picture

                              End If

               End If

End Sub
