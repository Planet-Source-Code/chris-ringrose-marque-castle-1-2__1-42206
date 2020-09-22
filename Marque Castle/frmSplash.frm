VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmSplash 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Marque Castle"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   HasDC           =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoading 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6330
      Top             =   3840
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Height          =   285
      Left            =   5880
      Picture         =   "frmSplash.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Launch Marque Castle"
      Top             =   4260
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Shape shpBoarder 
      BorderWidth     =   2
      Height          =   4950
      Left            =   15
      Top             =   15
      Width           =   6870
   End
   Begin WMPLibCtl.WindowsMediaPlayer medMidi 
      Height          =   2385
      Left            =   6900
      TabIndex        =   21
      Top             =   -300
      Visible         =   0   'False
      Width           =   2730
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
   End
   Begin VB.Label lblCheat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheats Enabled!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   1650
      TabIndex        =   20
      Top             =   765
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgBottomBoarder01 
      Height          =   465
      Left            =   0
      Picture         =   "frmSplash.frx":27B8
      ToolTipText     =   "http://www.gavannon.com/"
      Top             =   4485
      Width           =   4890
   End
   Begin VB.Image imgBottomBoarder02 
      Height          =   465
      Left            =   4860
      Picture         =   "frmSplash.frx":9EA8
      Stretch         =   -1  'True
      Top             =   4485
      Width           =   4425
   End
   Begin VB.Shape shpLoading 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   105
      Left            =   4320
      Top             =   3960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape shpLoadingTotal 
      BorderWidth     =   2
      Height          =   195
      Left            =   4290
      Shape           =   4  'Rounded Rectangle
      Top             =   3930
      Width           =   2385
   End
   Begin VB.Label lblStory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":A007
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   4320
      TabIndex        =   19
      Top             =   870
      Width           =   2265
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   3045
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   750
      Width           =   2505
   End
   Begin VB.Image imgPicture02 
      Height          =   240
      Left            =   3060
      Stretch         =   -1  'True
      ToolTipText     =   "Key on Cement"
      Top             =   1695
      Width           =   240
   End
   Begin VB.Image imgPicture01 
      Height          =   240
      Left            =   2745
      Stretch         =   -1  'True
      ToolTipText     =   "Key on Grass"
      Top             =   1695
      Width           =   240
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Used to Open Locked Blocks and Locked Doors (your objective)."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   2745
      TabIndex        =   18
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Scenarios"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   16
      Left            =   1215
      TabIndex        =   17
      ToolTipText     =   "Advanced"
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sound"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   15
      Left            =   1215
      TabIndex        =   16
      ToolTipText     =   "Option"
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Skins"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   14
      Left            =   1215
      TabIndex        =   15
      ToolTipText     =   "Option"
      Top             =   1530
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Death Mouse"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   13
      Left            =   1215
      TabIndex        =   14
      ToolTipText     =   "Adversary"
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Drone Mouse"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   12
      Left            =   1215
      TabIndex        =   13
      ToolTipText     =   "Adversary"
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bags of Cement"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   11
      Left            =   210
      TabIndex        =   12
      ToolTipText     =   "Item"
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bombs"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   10
      Left            =   210
      TabIndex        =   11
      ToolTipText     =   "Item"
      Top             =   2730
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Clock"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   9
      Left            =   210
      TabIndex        =   10
      ToolTipText     =   "Item"
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Metalic Boots"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   8
      Left            =   210
      TabIndex        =   9
      ToolTipText     =   "Item"
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Toggle Blocks"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   7
      Left            =   210
      TabIndex        =   8
      ToolTipText     =   "Obstacle"
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Doors"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   6
      Left            =   210
      TabIndex        =   7
      ToolTipText     =   "Objective"
      Top             =   2130
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Water"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   5
      Left            =   210
      TabIndex        =   6
      ToolTipText     =   "Obstacle"
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Blocks"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   4
      Left            =   210
      TabIndex        =   5
      ToolTipText     =   "Obstacle"
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Locked Blocks"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   3
      Left            =   210
      TabIndex        =   4
      ToolTipText     =   "Obstacle"
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Spikes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   2
      Left            =   210
      TabIndex        =   3
      ToolTipText     =   "Obstacle"
      Top             =   1530
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tiles"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   150
      Index           =   1
      Left            =   210
      TabIndex        =   2
      ToolTipText     =   "Object"
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label lblObjects 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Keys"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   150
      Index           =   0
      Left            =   210
      TabIndex        =   1
      ToolTipText     =   "Item"
      Top             =   1230
      Width           =   975
   End
   Begin VB.Image imgLogo 
      Height          =   1125
      Left            =   0
      Picture         =   "frmSplash.frx":A23B
      ToolTipText     =   "© 2003 Chris Ringrose"
      Top             =   0
      Width           =   3555
   End
   Begin VB.Image imgInfo 
      Height          =   2175
      Left            =   0
      Picture         =   "frmSplash.frx":B60F
      Top             =   1125
      Width           =   3960
   End
   Begin VB.Image imgBackground 
      Height          =   4995
      Left            =   0
      Picture         =   "frmSplash.frx":C3E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8565
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Marque Castle v1.2
'frmSplash
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

Dim strPassword As String




'  Code on Startup
Private Sub Form_Load()
  Dim i As Integer
  Dim sPicName As String

  On Local Error Resume Next
  
  'Initialize the random seed
  Randomize Timer
  
  'Starts to Play the Theme Song
  medMidi.settings.volume = 100
  medMidi.URL = App.Path & "\Marque Theme.mid"
  
  'Appropriately positions the objects on the form
  'The bottom Picture saying "Marque Castle"
  imgBottomBoarder01.Top = frmSplash.Height - imgBottomBoarder01.Height
  'The extension of the bottom Picture (Stretched)
  imgBottomBoarder02.Top = frmSplash.Height - imgBottomBoarder02.Height
  imgBottomBoarder02.Left = 0
  'The thin black Boarder arround the Form
  shpBoarder.Height = frmSplash.Height
  shpBoarder.Width = frmSplash.Width
  
  'Appropriately resizes the objects
  'The extension of the bottom Picture (Streches it)
  imgBottomBoarder02.Width = frmSplash.Width
  'The background Picture (Stretches it)
  imgBackground.Height = frmSplash.Height
  imgBackground.Width = frmSplash.Width
  
  'Shows the Form
  frmSplash.Show
  
  'Loads the custom settings
  'Sets the Value of strFile
  strFile = App.Path & "\CusSet.opt"
  
  'Sets the Value of fraPictures to store whether Reading or Writing
  frmMain.fraPictures.Caption = "Read"
  'Checks for Errors
  FileExists strFile
  
  'Assigns the Loaded Variables
  Input #1, strProperties(0), strProperties(1), strProperties(2), strProperties(3), _
  strProperties(4)
  'Decripts the Information
  For i = 0 To 4
    frmMain.lblStrProperties.Caption = i
    DecryptedCusSet "strProperties(" & frmMain.lblStrProperties.Caption & ")"
  Next
  
  'Sets the Saved Skin Directory
  strSkinDir = App.Path & "\skins\" & strProperties(0)
  'Puts the Skin Directory into the Change Skin Directory Text Box
  frmMain.txtSkinDir.Text = strProperties(0)

  'hide the Unregistered Label
  frmMain.lblRegistered.Visible = False

  'Loads all of the Items, Obstacles,and Textures
  For i = 0 To 38
    'Sets the Caption Value of the Pictures Frame for storage purposes
    sPicName = Format$(CStr(i), "000")
    
    'Loads the Picture into the appropriate Image slot for later retrieval
    frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & sPicName & ".bmp")
    If Err.Number Then
      Err.Clear
      frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & sPicName & ".gif")
    End If
    If Err.Number Then
      Err.Clear
      frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & sPicName & ".jpg")
    End If
  Next
  
  'Loads all of the Pictures of George
  For i = 91 To 95
    sPicName = "0" & CStr(i)
    frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & sPicName & ".bmp")
    If Err.Number Then
      Err.Clear
      frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & sPicName & ".gif")
    End If
    If Err.Number Then
      Err.Clear
      frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & sPicName & ".jpg")
    End If
  Next

  'Loads all of the Pictures for Norman
  For i = 891 To 895
    'Loads the Picture into the appropriate Image slot for later retrieval
    frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & CStr(i) & ".bmp")
    If Err.Number Then
      Err.Clear
      frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & CStr(i) & ".gif")
    End If
    If Err.Number Then
      Err.Clear
      frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & CStr(i) & ".jpg")
    End If
  Next

  'Loads all of the Pictures of all the Adversaries
  For i = 991 To 993
    'Loads the Picture into the appropriate Image slot for later retrieval
    frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & CStr(i) & ".bmp")
    If Err.Number Then
      Err.Clear
      frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & CStr(i) & ".gif")
    End If
    If Err.Number Then
      Err.Clear
      frmMain.pic(i).Picture = LoadPicture(strSkinDir & "\" & CStr(i) & ".jpg")
    End If
  Next
  
  'Loads the Picture for each Key (Grass and Cement)
  imgPicture01.Picture = frmMain.pic(2)
  imgPicture02.Picture = frmMain.pic(3)
  
  'Loads the Battle Areana character Pictures
  frmMain.imgPlayer(0).Picture = frmMain.pic(93).Picture
  frmMain.imgPlayer(1).Picture = frmMain.pic(893).Picture
  
  'Begins the fake "loading" bar
  tmrLoading.Enabled = True
  
  If Err.Number Then
    Err.Clear
    MsgBox "There was an error in loading one or more of the skin files." & vbNewLine & "Check your skins directory for all the files.", vbCritical, "Marque Castle - Loading Error"
  End If
End Sub

'  Gives a fake Loading Sequence
Private Sub tmrLoading_Timer()
  Dim i As Integer

     'Makes the Shape Visible
          shpLoading.Visible = True

     'Otherwise, make the Loading Shape longer
          For i = 1 To 150

               'Ensures that it isn't done Loading
                    If shpLoading.Width >= 2295 Then

                         shpLoading.Width = 2295
                         cmdEnter.Visible = True
                         shpLoading.Visible = False
                         shpLoadingTotal.Visible = False
                         tmrLoading.Enabled = False

                    End If

               shpLoading.Width = shpLoading.Width + 12
               Sleep 10
               DoEvents

          Next

     'Increases the Timer's speed (Interval)
          If tmrLoading.Interval > 0 Then

               tmrLoading.Interval = tmrLoading.Interval - 25

          End If

    If Me.Visible Then Me.SetFocus
End Sub


'  When you Move the Mouse over
Private Sub lblObjects_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer

     'Deselects all the Options
          For i = 0 To 16

               lblObjects(i).ForeColor = &HC0E0FF
          Next

     'Selects the Option selected
          lblObjects(Index).ForeColor = &H0&

     Select Case lblObjects(Index).Index

          'Keys
               Case 0

                    lblDescription.Caption = "Used to Open Locked Doors (your objective) and Locked Blocks."
                    imgPicture01.Picture = frmMain.pic(2)
                    imgPicture02.Picture = frmMain.pic(3)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Key on Grass"
                    imgPicture02.Visible = True: imgPicture02.ToolTipText = "Key on Cement"

          'Tiles
               Case 1

                    lblDescription.Caption = "Reverse all the Toggle Blocks on the Map.  (If one's On it becomes Off.  If one's Off it becomes On.)"
                    imgPicture01.Picture = frmMain.pic(4)
                    imgPicture02.Picture = frmMain.pic(5)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Tile on Grass"
                    imgPicture02.Visible = True: imgPicture02.ToolTipText = "Tile on Cement"

          'Spikes
               Case 2

                    lblDescription.Caption = "Unless George has the metallic Boots, he will die if he steps on these!"
                    imgPicture01.Picture = frmMain.pic(6)
                    imgPicture02.Picture = frmMain.pic(7)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Spikes on Grass"
                    imgPicture02.Visible = True: imgPicture02.ToolTipText = "Spikes on Cement"

          'Locked Blocks
               Case 3

                    lblDescription.Caption = "These block off George's path.  They require a Key to get past."
                    imgPicture01.Picture = frmMain.pic(8)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Locked Block"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Blocks
               Case 4

                    lblDescription.Caption = "Acts as an obstacle for George.  They can be blown up with a Bomb."
                    imgPicture01.Picture = frmMain.pic(9)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Block"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Water
               Case 5

                    lblDescription.Caption = "Unless George has a Bag of Cement, he will drown."
                    imgPicture01.Picture = frmMain.pic(13)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Water"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Doors
               Case 6

                    lblDescription.Caption = "This is George's objective in every Level.  They require a Key to enter."
                    imgPicture01.Picture = frmMain.pic(23)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Door"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Toggle Blocks
               Case 7

                    lblDescription.Caption = "When Off, George may walk upon these.  When On, he may not.  Their state is reversed by stepping on a 'Tile''"
                    imgPicture01.Picture = frmMain.pic(33)
                    imgPicture02.Picture = frmMain.pic(34)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Toggle Block (Off)"
                    imgPicture02.Visible = True: imgPicture02.ToolTipText = "Toggle Block (On)"

          'metallic Boots
               Case 8

                    lblDescription.Caption = "These allow George to walk on Spikes without getting hurt.  He may only carry one pair per Level."
                    imgPicture01.Picture = frmMain.pic(35)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "metallic Boots"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Clock
               Case 9

                    lblDescription.Caption = "This will reset the Timer back to 150.  George may only carry one per Level."
                    imgPicture01.Picture = frmMain.pic(36)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Clock"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Bombs
               Case 10

                    lblDescription.Caption = "Press SPACE BAR to use.  Will blow up Adversaries and Blocks.  George may only carry one at a time."
                    imgPicture01.Picture = frmMain.pic(37)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Bomb"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Cement Bags
               Case 11

                    lblDescription.Caption = "When in George's possesion, Water that he walks on will become solid ground."
                    imgPicture01.Picture = frmMain.pic(38)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Cement Bags"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Drone Mouse
               Case 12

                    lblDescription.Caption = "A simplistic Adversary that follows boundaries.  It will explode when it dies...."
                    imgPicture01.Picture = frmMain.pic(991)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Drone Mouse"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Death Mouse
               Case 13

                    lblDescription.Caption = "A mildly intelligent Adversary that stalks George throughout the Level."
                    imgPicture01.Picture = frmMain.pic(992)
                    imgPicture01.Visible = True: imgPicture01.ToolTipText = "Death Mouse"
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Skins
               Case 14

                    lblDescription.Caption = "The style or group of pictures that Marque Castle uses throughout the Game."
                    imgPicture01.Visible = False: imgPicture01.ToolTipText = ""
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Sound
               Case 15

                    lblDescription.Caption = "Involves both Music and Sound Effects.  Both may be disabled by your command."
                    imgPicture01.Visible = False: imgPicture01.ToolTipText = ""
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

          'Custom Scenarios
               Case 16

                    lblDescription.Caption = "Custom Levels can be made with the Scenario Creation Artist!"
                    imgPicture01.Visible = False: imgPicture01.ToolTipText = ""
                    imgPicture02.Visible = False: imgPicture02.ToolTipText = ""

     End Select

End Sub


'Starts Marque Castle!
Private Sub cmdEnter_Click()
  Beep
    
  'Shows the Main Form
  frmMain.Show
  
  'Hides the Splash Screen
  frmSplash.Hide
  
  If lblCheat.Visible = True Then frmMain.lblCheater.Visible = True
  
  'Messages
  frmMain.tmrNewGame = True

  Do
    Sleep 6
    DoEvents
    medMidi.settings.volume = medMidi.settings.volume - 1
    Sleep 7
    DoEvents
  Loop Until medMidi.settings.volume = 35
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyG Then
          strPassword = strPassword & "g"
     ElseIf KeyCode = vbKeyA Then
          strPassword = strPassword & "a"
     ElseIf KeyCode = vbKeyV Then
          strPassword = strPassword & "v"
     ElseIf KeyCode = vbKeyN Then
          strPassword = strPassword & "n"
     ElseIf KeyCode = vbKeyO Then
          strPassword = strPassword & "o"
     ElseIf KeyCode = vbKey1 Then
          strPassword = strPassword & "1"
     ElseIf KeyCode = vbKey2 Then
          strPassword = strPassword & "2"
     ElseIf KeyCode = vbKey3 Then
          strPassword = strPassword & "3"
     Else
          strPassword = ""
     End If

     If strPassword = "gavannon123" Then
          lblCheat.Visible = True
          intLivesNum = 5
          frmMain.lblLives = intLivesNum
          frmMain.tmrDroneAI.Interval = 500
          frmMain.tmrDeathAI.Interval = 500
          frmMain.tmrTimer.Interval = 1500
          'frmMain.lblSource.Visible = True
          frmMain.lblSpeed.Visible = True
     End If

End Sub
