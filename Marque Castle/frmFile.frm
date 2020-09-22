VERSION 5.00
Begin VB.Form frmFile 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data - Unknown"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000080FF&
   HasDC           =   0   'False
   Icon            =   "frmFile.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4140
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4140
      Begin VB.TextBox txtLoading 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   990
         TabIndex        =   9
         Text            =   "L O A D I N G . . . ."
         Top             =   2640
         Width           =   2100
      End
      Begin VB.Shape shpLoading 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         Height          =   315
         Left            =   975
         Top             =   2625
         Width           =   2130
      End
   End
   Begin VB.FileListBox filFileListBox 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2400
      Left            =   150
      Pattern         =   "*.lvl"
      TabIndex        =   1
      Top             =   450
      Width           =   1890
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2220
      TabIndex        =   0
      Text            =   "*.cus"
      Top             =   120
      Width           =   1815
   End
   Begin VB.DriveListBox drvDrive 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2220
      TabIndex        =   2
      Top             =   450
      Width           =   1815
   End
   Begin VB.DirListBox dirDirectory 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2040
      Left            =   2220
      TabIndex        =   3
      Top             =   780
      Width           =   1815
   End
   Begin VB.CommandButton cmdSaveLoad 
      Appearance      =   0  'Flat
      Caption         =   "Save/Load"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      Picture         =   "frmFile.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3075
      Width           =   765
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3225
      Picture         =   "frmFile.frx":044F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3075
      Width           =   765
   End
   Begin VB.Label lblFileNameTitle 
      BackColor       =   &H0080C0FF&
      Caption         =   "*.cus"
      Height          =   285
      Left            =   2220
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   150
      TabIndex        =   7
      Top             =   3075
      Width           =   1890
   End
   Begin VB.Label lblFileName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   225
      TabIndex        =   6
      Top             =   150
      Width           =   1815
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Marque Castle v1.2
'frmFile
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


'  Cancel Loading or Saving>>
Private Sub cmdCancel_Click()
  Unload frmFile
End Sub
     

'  Saving in either Scenario Creation Artist or Marque Castle
Private Sub cmdSaveLoad_Click()
  Dim i As Integer
     
     'No File Name entered
          If txtFileName.Text = "" Or txtFileName.Text = "*.cus" Then
               cmdSaveLoad.Enabled = False
               txtFileName.FontBold = True

     'A proper File Name entered
          Else

               'You're Saving....
                     If cmdSaveLoad.Caption = "Save" Then
     
                         '....In Scenario Creation Artist (Your Scenario)
                              If lblTitle.Caption = "Scenario Creation Artist" Then
     
                                 SaveCreation
     
                         '....In Marque Castle (Your Progress)
                             ElseIf lblTitle.Caption = "Marque Castle" Then
     
     
                             End If
     
                 'You're Loading....
                     Else
     
                         '....In Scenario Creation Artist (Your Scenario)
                             If lblTitle.Caption = "Scenario Creation Artist" Then

                                 LoadCreation
     
                         '....In Marque Castle (A Custom Scenario)
                              ElseIf lblTitle.Caption = "Marque Castle" And lblFileName.Caption = "Scenario to Load:" Then

                                   'Warns you that this will start a new game before actually doing so
                                        If MsgBox("This will Quit any current Game Opened", vbOKCancel, "Marque Castle") = vbOK Then
                                             For i = 0 To 399
                                                  frmMain.imgMap(i).Visible = False
                                             Next
                                             'The name of the file to load:
                                             strFile = dirDirectory & "\" & frmFile.txtFileName.Text
                                             If frmSplash.lblCheat.Visible = False Then
                                                  intLivesNum = 5
                                             Else
                                                  intLivesNum = 10
                                             End If
                                             intKeysNum = 0

                                             ' Hides frmFile
                                             frmFile.Visible = False
                                             frmMain.Enabled = True
'                                             frmMain.Show

                                             'Loads the Level
                                                  strBestTimesDir = dirDirectory
                                                  LoadLevel

                                             frmMain.fraPaused.Visible = False
                                             frmMain.fraDefeat.Visible = False
                                             'Resets the Score
                                                  dblScore = 0
                                                  frmMain.lblScore.Caption = "0000000000"

                                             'Loads the Pictures into their places
                                                  frmMain.imgGeorge.Picture = frmMain.pic(93).Picture
                                                  frmMain.imgKey.Picture = frmMain.pic(3).Picture
                                                  frmMain.imgCement.Picture = frmMain.pic(38).Picture
                                                  frmMain.imgBoots.Picture = frmMain.pic(35).Picture
                                                  frmMain.imgBoots.Visible = False
                                                  frmMain.imgClock.Picture = frmMain.pic(36).Picture
                                                  frmMain.imgClock.Visible = False
                                                  frmMain.imgBomb.Picture = frmMain.pic(37).Picture
                                                  frmMain.imgBomb.Visible = False
                                                  frmMain.imgExplosion.Picture = LoadPicture(strSkinDir & "\Explosion.gif")
                                                  frmMain.imgExplosionSmall.Picture = LoadPicture(strSkinDir & "\Explosion.gif")
                                                  frmMain.imgExplosion.Visible = False
                                             
                                             'Unloads frmFile
                                                  Unload frmFile

                                        End If
     
                             End If
     
                 End If

     End If

End Sub
     
     
Private Sub dirDirectory_Change()
     On Error Resume Next
     'Upon changing the directory, update the File List
          filFileListBox.Path = dirDirectory.Path
     
End Sub
     
     
Private Sub drvDrive_Change()
     On Error Resume Next
         'Upon changing the drive
             dirDirectory.Path = drvDrive.Drive
     
End Sub
     
     
Private Sub filFileListBox_Click()
         On Error Resume Next
         'Upon selecting a file
             txtFileName.FontBold = False
             txtFileName.Text = filFileListBox.FileName
             lblFileNameTitle.Caption = filFileListBox.FileName
             cmdSaveLoad.Enabled = True
     
End Sub
     
     
Private Sub filFileListBox_DblClick()
  Dim i As Integer

     'No File Name entered
          If txtFileName.Text = "" Or txtFileName.Text = "*.cus" Then
               cmdSaveLoad.Enabled = False
               txtFileName.FontBold = True

     'A proper File Name entered
          Else

               'You're Saving....
                     If cmdSaveLoad.Caption = "Save" Then
     
                         '....In Scenario Creation Artist (Your Scenario)
                              If lblTitle.Caption = "Scenario Creation Artist" Then
     
                                 SaveCreation
     
                         '....In Marque Castle (Your Progress)
                             ElseIf lblTitle.Caption = "Marque Castle" Then
     
     
                             End If
     
                 'You're Loading....
                     Else
     
                         '....In Scenario Creation Artist (Your Scenario)
                             If lblTitle.Caption = "Scenario Creation Artist" Then

                                 LoadCreation
     
                         '....In Marque Castle (A Custom Scenario)
                              ElseIf lblTitle.Caption = "Marque Castle" And lblFileName.Caption = "Scenario to Load:" Then

                                   'Warns you that this will start a new game before actually doing so
                                        If MsgBox("This will Quit any current Game Opened", vbOKCancel, "Marque Castle") = vbOK Then
                                             For i = 0 To 399
                                                  frmMain.imgMap(i).Visible = False
                                             Next
                                             'The name of the file to load:
                                             strFile = dirDirectory & "\" & frmFile.txtFileName.Text
                                             If frmSplash.lblCheat.Visible = False Then
                                                  intLivesNum = 5
                                             Else
                                                  intLivesNum = 10
                                             End If
                                             intKeysNum = 0

                                             ' Hides frmFile
                                             frmFile.Visible = False
                                             frmMain.Enabled = True
'                                             frmMain.Show

                                             'Loads the Level
                                                strBestTimesDir = dirDirectory
                                                LoadLevel

                                             frmMain.fraPaused.Visible = False
                                             frmMain.fraDefeat.Visible = False
                                             'Resets the Score
                                                  dblScore = 0
                                                  frmMain.lblScore.Caption = "0000000000"

                                             'Loads the Pictures into their places
                                                  frmMain.imgGeorge.Picture = frmMain.pic(93).Picture
                                                  frmMain.imgKey.Picture = frmMain.pic(3).Picture
                                                  frmMain.imgCement.Picture = frmMain.pic(38).Picture
                                                  frmMain.imgBoots.Picture = frmMain.pic(35).Picture
                                                  frmMain.imgBoots.Visible = False
                                                  frmMain.imgClock.Picture = frmMain.pic(36).Picture
                                                  frmMain.imgClock.Visible = False
                                                  frmMain.imgBomb.Picture = frmMain.pic(37).Picture
                                                  frmMain.imgBomb.Visible = False
                                                  frmMain.imgExplosion.Picture = LoadPicture(strSkinDir & "\Explosion.gif")
                                                  frmMain.imgExplosionSmall.Picture = LoadPicture(strSkinDir & "\Explosion.gif")
                                                  frmMain.imgExplosion.Visible = False
                                             
                                             'Unloads frmFile
                                                  Unload frmFile

                                        End If
     
                             End If
     
                 End If

     End If

End Sub


Private Sub Form_Load()
    dirDirectory.Path = App.Path
    On Error Resume Next
    dirDirectory.Path = App.Path & "\My Scenarios"
End Sub


Private Sub Form_Unload(Cancel As Integer)

     'If you're creating Custom Game
          If lblTitle.Caption = "Scenario Creation Artist" Then

               'Enables the Scenario Creation Artist
                    frmCreate.Enabled = True

     'You're Loading in Marque Castle
          Else

               'Enables the Main Form
                    frmMain.Enabled = True

               'Shows the Main Form
'                    frmMain.Show

          End If

End Sub


Private Sub txtFileName_Change()
  Dim intErrors As Integer
  Dim i As Integer

    If Len(txtFileName.Text) < 5 Then intErrors = 1
    If Len(txtFileName.Text) > 3 Then
        If Mid$(txtFileName, Len(txtFileName.Text) - 3, 4) <> ".cus" Then intErrors = 1
    Else
        intErrors = 1
    End If
    For i = 1 To Len(txtFileName.Text)
        If Mid$(txtFileName.Text, i, 1) = "\" Then intErrors = 1
        If Mid$(txtFileName.Text, i, 1) = "/" Then intErrors = 1
        If Mid$(txtFileName.Text, i, 1) = ":" Then intErrors = 1
        If Mid$(txtFileName.Text, i, 1) = "*" Then intErrors = 1
        If Mid$(txtFileName.Text, i, 1) = "?" Then intErrors = 1
        If Mid$(txtFileName.Text, i, 1) = "<" Then intErrors = 1
        If Mid$(txtFileName.Text, i, 1) = ">" Then intErrors = 1
        If Mid$(txtFileName.Text, i, 1) = "|" Then intErrors = 1
        If Mid$(txtFileName.Text, i, 1) = ";" Then intErrors = 1
    Next
    If intErrors = 0 Then
        cmdSaveLoad.Enabled = True
    Else
        cmdSaveLoad.Enabled = False
    End If
End Sub

Private Sub txtFileName_GotFocus()
    'Sets the Text to Normal (not Bold) if
        txtFileName.FontBold = False
End Sub
