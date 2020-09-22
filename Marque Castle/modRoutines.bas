Attribute VB_Name = "modRoutines"


'Marque Castle v1.2
'modRoutines
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




'Tracks the number of these objects
     Dim intKey As Integer
     Dim intDoor As Integer
     Dim intGeorge As Integer
     Dim intNorman As Integer
     Dim intToggleBlock As Integer
     Dim intBomb As Integer
     Dim intClock As Integer

'The Error Description upon Saving
     Dim strError As String



Public Function RndBetween(ByRef MinNum As Integer, ByRef MaxNum As Integer) As Integer
  RndBetween = Int((MaxNum - MinNum + 1) * Rnd + MinNum)
End Function


'         <<Saves a Level in the Level Creation Artist>>
Public Sub SaveCreation()
  Dim i As Integer

     'Shows the Loading Frame (Black Screen)

          frmFile.fraLoading.Visible = True

     'Adds the File Extension in the File Name if not present (.cus)

          If Right(frmFile.txtFileName.Text, 4) <> ".cus" Then frmFile.txtFileName.Text = frmFile.txtFileName.Text & ".cus"

     'Assigns a Name to the Level if not present

          If frmCreate.txtLevelTitle.Text = "" Then frmCreate.txtLevelTitle.Text = "Untitled Level"

     'Checks the File Data before Saving

          'Reset the following counters

               intKey = 0
               intDoor = 0
               intGeorge = 0
               intNorman = 0
               intToggleBlock = 0
               intBomb = 0
               intClock = 0
               strData = ""

          'Counts the number of the following objects & Loads Data into strData

               For i = 0 To 399

                    'Checks for Keys
                         If strCell(i) = "002" Or strCell(i) = "003" Then intKey = intKey + 1

                    'Checks for Doors
                         If strCell(i) = "023" Then intDoor = intDoor + 1

                    'Checks for George
                         If strCell(i) = "093" Then intGeorge = intGeorge + 1

                    'Checks for Norman
                         If strCell(i) = "893" Then intNorman = intNorman + 1

                    'Checks for Toggle Blocks
                         If strCell(i) = "033" Or strCell(i) = "034" Then intToggleBlock = intToggleBlock + 1

                    'Checks for Bombs
                         If strCell(i) = "037" Then intBomb = intBomb + 1

                    'Checks for Clocks
                         If strCell(i) = "036" Then intClock = intClock + 1

                    'Loads the data into strData
                         strData = strData & strCell(i)

               Next

     'Loads the following Variables into the appropriate place

               'The Level Number
                    strLevel = frmCreate.cboLevelNumber.Text

               'What the Next Level is (If there is one)

                    'No following Level (Last Level)
                         If frmCreate.chkLastLevel.Value = 1 Then
                            strNextLevel = "Last"

                    'There is a Next Level
                         ElseIf frmCreate.chkLastLevel.Value = 0 Then
                            strNextLevel = frmCreate.txtNextLevel.Text
                        End If

               'The Title of the Level
                    strLevelTitle = frmCreate.txtLevelTitle.Text

               'The Message to Gamer on Startup
                    strMessage = frmCreate.txtMessage.Text

               'The Authors Name or Pseudonym
                    strAuthor = frmCreate.txtAuthor.Text

               'The ground below your feet

                    'The ground is Grass
                         If frmCreate.lblGrass.BorderStyle = 1 Then
                              strGround = "000"

                    'The ground is Cement
                         Else
                              strGround = "001"
                         End If

     'Displays Error Message on Error

          'Resets the Error Checker's Variables and Attributes
               frmCreate.fraError.Visible = False
               frmCreate.cboErrors.Clear

          'No Key
               If intKey = 0 Then
                    frmCreate.fraError.Visible = True
                    strError = "No Key on map"
                    frmCreate.cboErrors.AddItem (strError)
               End If

          'More than one Door
               If intDoor > 1 Then
                    frmCreate.fraError.Visible = True
                    strError = "More than one Door"
                    frmCreate.cboErrors.AddItem (strError)
               End If

         'No Door
               If intDoor = 0 Then
                    frmCreate.fraError.Visible = True
                    strError = "No Door on map"
                    frmCreate.cboErrors.AddItem (strError)
               End If

          'More than one George Starting Position
               If intGeorge > 1 Then
                    frmCreate.fraError.Visible = True
                    strError = "More than one Starting Possition (George)"
                    frmCreate.cboErrors.AddItem (strError)
               End If

          'More than one Norman Starting Position
               If intNorman > 1 Then
                    frmCreate.fraError.Visible = True
                    strError = "More than one Starting Possition (Norman)"
                    frmCreate.cboErrors.AddItem (strError)
               End If

          'No Starting Position (George)
               If intGeorge = 0 Then
                    frmCreate.fraError.Visible = True
                    strError = "No Starting Possition (George) on map"
                    frmCreate.cboErrors.AddItem (strError)
               End If

          'More than a two-hundered and fifty Toggle Blocks
               If intToggleBlock > 250 Then
                    frmCreate.fraError.Visible = True
                    strError = "More than 250 Toggle Blocks (On or Off)"
                    frmCreate.cboErrors.AddItem (strError)
               End If

          'Not all boxes are used
               If Len(strData) < 1200 Then
                    frmCreate.fraError.Visible = True
                    strError = "Not all boxes are used"
                    frmCreate.cboErrors.AddItem (strError)
               End If

          'Level Number unspecifed
               If strLevel = "" Then
                    frmCreate.fraError.Visible = True
                    strError = "Level Number not specified"
                    frmCreate.cboErrors.AddItem (strError)
               End If

          'If not Last Level, Next Level unspecified....
               If frmCreate.chkLastLevel.Value = 0 And Len(strNextLevel) < 5 Then
                    frmCreate.fraError.Visible = True
                    strError = "Following level unspecified"
                    frmCreate.cboErrors.AddItem (strError)
               End If

          'Leave if any Errors, and Disable the following objects
               If frmCreate.fraError.Visible = True Then
                    frmCreate.chkGrid.Enabled = False
                    frmCreate.chkPaintFill.Enabled = False
                    frmCreate.cmdUndo.Enabled = False
                    frmCreate.lblGrass.Enabled = False
                    frmCreate.lblCement.Enabled = False
                    frmCreate.cmdSave.Enabled = False
                    frmCreate.cmdLoad.Enabled = False
                    frmCreate.txtLevelTitle.Enabled = False
                    frmCreate.cboLevelNumber.Enabled = False
                    frmCreate.chkLastLevel.Enabled = False
                    frmCreate.txtNextLevel.Enabled = False
                    frmCreate.txtMessage.Enabled = False
                    frmCreate.txtAuthor.Enabled = False
                    frmCreate.txtMessage.Enabled = False
                    frmCreate.txtAuthor.Enabled = False
                    frmCreate.cmdDetails.Enabled = False
                    frmCreate.cmdDetails.Caption = "Hide"
                    Do While frmCreate.fraLevelManager.Height < 3375
                         frmCreate.fraLevelManager.Height = frmCreate.fraLevelManager.Height + 95
                         frmCreate.fraLevelManager.Top = frmCreate.fraLevelManager.Top - 95
                    Loop
                    Unload frmFile
                    Exit Sub
               End If
        
     'No Errors with Data:

          'Closes all Files Opened (If any)
               Close

          'Creates the File (Doesn't exist) or Writes on top of it (Does exist)
               strFile = frmFile.dirDirectory & "\" & frmFile.txtFileName.Text
               Open strFile For Output As #1

          'Inputs the Variable Data into the File
               Write #1, strLevel, strNextLevel, strLevelTitle, strMessage, strData, strAuthor, strGround

     'Renames the Form accordinly
          frmCreate.Caption = "Scenario Creation Artist - " & strLevelTitle
          frmCreate.txtFilePath.Text = strFile
          frmCreate.txtFilePath.SelStart = Len(frmCreate.txtFilePath.Text)

     'Hides the Loading Frame (Black screen)
          frmFile.fraLoading.Visible = False

     'Creates the Top Times Data Storage
          'Closes all Files Opened (If any)
               Close

          'Creates the File (Doesn't exist) or Writes on top of it (Does exist)
               strFile = frmFile.dirDirectory & "\" & frmCreate.lblBestTimesDir.Caption
               Open strFile For Output As #1

          'Inputs the Variable Data into the File
               For i = 0 To 4
                    Write #1, "Player Name", "000"
               Next

     'Closes the frmFile Form
          Unload frmFile

End Sub


'   <<Loads a Level in the Level Creation Artist>>
Public Sub LoadCreation()
  Dim i As Integer
  Dim BlocksNum As Integer
  Dim p As Integer

    On Error Resume Next
     'Shows the Loading Frame (Black screen)
          frmFile.fraLoading.Visible = True

     'Fix up the File Name before Loading

          'Adds the extension if not present
               If Right(frmFile.txtFileName.Text, 4) <> ".cus" Then frmFile.txtFileName.Text = frmFile.txtFileName.Text & ".cus"

     'Opens the File

          'Determines what to Open and Opens it
               strFile = frmFile.dirDirectory & "\" & frmFile.txtFileName.Text

          'Sets the Value of fraPictures to store whether Reading or Writing
               frmMain.fraPictures.Caption = "Read"
          'Checks for Errors
               FileExists strFile
          
          'Inputs the information into the appropriate Variables
               Input #1, strLevel, strNextLevel, strLevelTitle, strMessage, strData, strAuthor, strGround

     'Loads the following Variables into the appropriate places

          'The Level Number
               frmCreate.cboLevelNumber.Text = strLevel

          'What the Next Level is (If there is one)
     
               'No following Level (Last Level)
                    If strNextLevel = "Last" Then
                         frmCreate.chkLastLevel.Value = 1
                         frmCreate.txtNextLevel.Enabled = False
                         frmCreate.txtNextLevel.Text = "*.cus"

               'There is a Next Level
                    Else
                         frmCreate.chkLastLevel.Value = 0
                         frmCreate.txtNextLevel.Enabled = True
                         frmCreate.txtNextLevel.Text = strNextLevel
                    End If

          'The Title of the Level
               frmCreate.txtLevelTitle.Text = strLevelTitle

          'The Message to the Gamer on Startup
               frmCreate.txtMessage.Text = strMessage

         ' Checks if it's the secret level
            If Mid$(strData, 1, 963) = "011011011011011011011023011011011012012012012012012012012012011003011001009000009009009000011011001001001001001001001012011009011001009000001037001000000011001001001001001001001012011991001001009000001001001000000011011001001001001001001012011009009009001000001001001001001007007001001001001001001012011000000000000000001001001001007007007007001001001001001012011000000000000000001001001007007007007007007001033033033012011000000034034034034000000000006006007007001001033001001012011000034033033033009034000000000000001001001001033001012005011000034033003033009034000000000000001001001001033001034093011000034033033033009034000000000000001001001001033001034893011000000034034034034000000000000000001001001001033001012005011000000000000000000000000000006006007007001001033001001012011013000013013000000000000006006006007007007001033033033012011000013013013013000000000000006006007007001001001009009012011013013013013013013000000000000006007001001001001009037012011" Then
                BlocksNum = 0
                p = 1
                For i = 0 To 399
                    If Mid$(strData, p, 3) = "009" Then BlocksNum = BlocksNum + 1
                    p = p + 3
                    frmCreate.imgMap(i).Enabled = False
                Next
                If BlocksNum = 20 Then
                    MsgBox "Beat this level leaving only four blocks behind, with 110 sec. remaining ... " & vbNewLine & " ... to learn a secret ...", vbInformation, "Marque Castle"
                Else
                    MsgBox "Hmmm ... this doesn't look the same somehow ...", vbInformation, "Marque Castle"
                End If
                frmCreate.cmdSave.Enabled = False
                frmCreate.cmdDetails.Enabled = False
            Else
                For i = 0 To 399
                    frmCreate.imgMap(i).Enabled = True
                Next
                frmCreate.cmdSave.Enabled = True
                frmCreate.cmdDetails.Enabled = True
            End If

          'Loads the Pictures into their places (On the Map)
               p = 1
               For i = 0 To 399
                    strCell(i) = Mid$(strData, p, 3)
                    frmCreate.imgMap(i).Picture = frmMain.pic(CInt(Mid(strData, p, 3))).Picture
                    p = p + 3
                    DoEvents
               Next

          'Loads the Author Name
               frmCreate.txtAuthor.Text = strAuthor

          'Loads the ground below your feet

               'The ground is Grass
                    If strGround = "000" Then
                         frmCreate.lblGrass.BorderStyle = 1
                         frmCreate.lblCement.BorderStyle = 0

               'The ground is Cement
                    Else
                         frmCreate.lblGrass.BorderStyle = 0
                         frmCreate.lblCement.BorderStyle = 1
                    End If

     'Renames the form accordinly
          frmCreate.Caption = "Scenario Creation Artist - " & strLevelTitle
          frmCreate.txtFilePath.Text = strFile
          frmCreate.txtFilePath.SelStart = Len(frmCreate.txtFilePath.Text)


     'Closes the Load Form
          Unload frmFile

     'Hides the Loading Frame (Black screen)
          frmFile.fraLoading.Visible = False

     'Disables the Main Form
          frmMain.Enabled = False

     'Enables the Scenario Creation Form
          frmCreate.Enabled = True
          frmCreate.Show

End Sub


'George dies
Public Sub Defeat()

     'Stops the Music
          frmSplash.medMidi.URL = ""

     'Stops the enimies
          frmMain.tmrDroneAI.Enabled = False
          frmMain.tmrDeathAI.Enabled = False

     'No Game is in progress
          booPlaying = False

     'Disables the Timer
          frmMain.tmrTimer.Enabled = False

     'The Game is not paused
          frmMain.fraPaused.Visible = False

     'Shows Dead Picture
          frmMain.imgMap(intPos).Picture = frmMain.pic(95).Picture

     'Displays Defeat Title
          frmMain.fraDefeat.Visible = True

     'Changes the Menu
          frmMain.mnuQuitGame.Enabled = True
          frmMain.mnuSaveGame.Enabled = False
          frmMain.mnuSaveGameAs.Enabled = False
          frmMain.mnuPauseGame.Enabled = False
          frmMain.mnuRestartLevel.Enabled = True

     'Subtracts 100 Points from Score
          If dblScore > 100 Then
               dblScore = dblScore - 100
          Else
               dblScore = 0
          End If
          ScoreUpdate

     'Checks for a Game Over
          If intLivesNum > 0 Then

               'Subtracts a Life
                    intLivesNum = intLivesNum - 1
                    frmMain.lblLives.Caption = intLivesNum

        Else

            MsgBox "Sorry dude, Game Over!", vbCritical, "Marque Castle"

            'Quits the Game
                QuitGame

          End If

End Sub


'  A Level is being Loaded
Public Sub LoadLevel()
  Dim intPauseAmm As Integer
  Dim BlocksNum As Integer
  Dim p As Integer
  Dim i As Integer

  On Error Resume Next


  With frmMain
    .fraPaused.Visible = False
    .fraDefeat.Visible = False
  
    .tmrNewGame.Enabled = False
    .lblItemInfo.Caption = ""
  
    'Checks if Music is Checked off
    If .mnuMusic.Checked = True Then
      'Randomly Plays a MIDI Song
      frmSplash.medMidi.URL = App.Path & "\Music(" & CStr(RndBetween(0, 4)) & ").mid"
    End If

    'Shows the appropriate stats
    .fraLevelInfo.Visible = True
    .fraItems.Visible = True
    .fraTimer.Visible = True

    'Shows the Loading Label
    .fraLoading.Visible = True
  
    'Sets the File Path Label to strFile
    .lblFilePath.Caption = strFile
  
    'Disables the Clock Timer
    .tmrTimer.Enabled = False
  
    'Resets the Clock
    .lblTimer.Caption = "150"
  
    'Closes all Files Opened (If any)
    Close
  
    'Sets the Value of fraPictures to store whether Reading or Writing
    .fraPictures.Caption = "Read"
  
    'Checks for Errors
    FileExists strFile
    strLevelFile = strFile
  
    'Loads the level data into the appropriate variables
    Input #1, strLevel, strNextLevel, strLevelTitle, strMessage, strData, strAuthor, strDefaultGround
  
    'The Title of the Level
    .lblLevelTitle.Caption = strLevelTitle
  
    'Resets the Drone Possision
    intDronePos = -1
  
    'Checks if you're in the secret level
    If Mid$(strData, 1, 963) = "011011011011011011011023011011011012012012012012012012012012011003011001009000009009009000011011001001001001001001001012011009011001009000001037001000000011001001001001001001001012011991001001009000001001001000000011011001001001001001001012011009009009001000001001001001001007007001001001001001001012011000000000000000001001001001007007007007001001001001001012011000000000000000001001001007007007007007007001033033033012011000000034034034034000000000006006007007001001033001001012011000034033033033009034000000000000001001001001033001012005011000034033003033009034000000000000001001001001033001034093011000034033033033009034000000000000001001001001033001034893011000000034034034034000000000000000001001001001033001012005011000000000000000000000000000006006007007001001033001001012011013000013013000000000000006006006007007007001033033033012011000013013013013000000000000006006007007001001001009009012011013013013013013013000000000000006007001001001001009037012011" Then
      blnSecret = True
      BlocksNum = 0
      p = 1
      For i = 0 To 399
        If Mid$(strData, p, 3) = "009" Then BlocksNum = BlocksNum + 1
        p = p + 3
        frmCreate.imgMap(i).Enabled = False
      Next
      If BlocksNum = 20 Then
        blnValidSecret = True
      Else
        MsgBox "Hmmm ... ", vbInformation, "Marque Castle"
        blnValidSecret = False
      End If
    Else
      blnSecret = False
      blnValidSecret = False
    End If
  
    'Loads the Pictures into their place on the map
    p = 1
    intDronePos = -1
    intDeathPos = -1
    intNormanPosition = -1
    For i = 0 To 399
      .imgMap(i).Picture = .pic(CInt(Mid(strData, p, 3))).Picture
      .imgMap(i).Visible = True
      strCell(i) = Mid$(strData, p, 3)
      If Mid$(strData, p, 3) = "093" Then intPos = i 'George
      If Mid$(strData, p, 3) = "991" Then intDronePos = i 'Drone Mouse
      If Mid$(strData, p, 3) = "992" Then intDeathPos = i 'Death Mouse
      If Mid$(strData, p, 3) = "893" Then intNormanPosition = i 'Norman
      p = p + 3
    Next
    
    'Sets strGround
    strGround = strDefaultGround
    strNormanGround = strDefaultGround
    strDeathGround = strDefaultGround
    
    'Renames the Form accordinly
    .Caption = "Marque Castle - " & strLevelTitle
    
    'Hides the Boots, Clock, and the Bomb
    If .lblCheater.Visible = False Then .imgBoots.Visible = False
    .imgClock.Visible = False
    If .lblCheater.Visible = False Then .imgBomb.Visible = False
    If .lblCheater.Visible = True Then .imgBoots.Visible = True
    If .lblCheater.Visible = True Then .imgBomb.Visible = True
  
    'Loads the Top 5 scores
    If blnSecret = False Then
      'Determines what to Open and Opens it
      If Right$(strLevelFile, 3) = "lvl" Then strFile = App.Path & "\Scenarios\Lvl" & strLevel & ".bt"
      If Right$(strLevelFile, 3) = "cus" Then strFile = strBestTimesDir & "\Lvl" & strLevel & ".bt"

      .lblCreator.Visible = IIf(Mid$(strLevelFile, Len(strLevelFile) - 2, 3) = "cus", True, False)
      
      'Sets the Value of fraPictures to store whether Reading or Writing
      .fraPictures.Caption = "Read"
      'Checks for Errors
      FileExists strFile
      
      'Loads the data
      For i = 0 To 4
        Input #1, strTopNames(i), intTopTimes(i)
      Next
      
      'Loads the data into the charts
      For i = 0 To 4
        .lblTopName(i).Caption = strTopNames(i)
        .lblTopScore(i).Caption = intTopTimes(i)
      Next
      .fraHighScore.Caption = "Level" & strLevel
      .mnuBestTimes.Enabled = True
    Else
      .mnuBestTimes.Enabled = False
    End If
  
    'Resets the Menu Properties acordingly
    .mnuQuitGame.Enabled = True
    .mnuSaveGame.Enabled = False
    .mnuSaveGameAs.Enabled = False
    .mnuPauseGame.Enabled = False
    .mnuRestartLevel.Enabled = False
    .lblCreator.Caption = strAuthor
    
    'Resets certain Properties
    intKeysNum = 0
    .lblKeysNum = intKeysNum
    .mnuPauseGame.Caption = "&Pause Game"
    .lblEnd(0).Visible = False
    .lblEnd(1).Visible = False
    .fraMessage.Visible = False
    .cmdBegin02.Visible = False
    
    'Hides the Loading Label
    .fraLoading.Visible = False
  
    'The message to the Gamer on Startup
    'There is a Message
    If Mid$(strMessage, 1, 1) <> "<" And strMessage <> "" Then
      'Shows the Message and Begin Button
      .fraMessage.Visible = True
      .cmdBegin.Enabled = False
      intPauseAmm = 100 - (Len(strMessage) \ 2)
      .txtMessage.Text = ""
      .lblLevelNum.Caption = "''Level " & strLevel & "''"
      i = 1
      Do While i <= Len(strMessage) And .fraMessage.Visible = True
        If i >= Len(strMessage) \ 2 Then .cmdBegin.Enabled = True
        If .cmdBegin.Enabled = True Then .cmdBegin.SetFocus
        .txtMessage.Text = .txtMessage.Text & Mid$(strMessage, i, 1)
        If intPauseAmm > 0 Then Sleep intPauseAmm
        DoEvents
        i = i + 1
      Loop
    'There isn't a Message
    Else
      'Shows the other Begin Button
      .cmdBegin02.Visible = True
    End If
    
    .tmrWatchTime.Enabled = True
  End With
End Sub


'Moves George
Public Sub MoveGeorge(Direction As Integer, Picture As Integer, George As Boolean)
  Dim p As Integer
  Dim i As Integer
  Dim r As Integer

  On Local Error Resume Next

     'If George is moving
          If George = True Then

               'Either <Grass> <Cement> or <Toggle Block OFF>
                    If CInt(strCell(intPos + Direction)) < 2 Or CInt(strCell(intPos + Direction)) = 33 Then

                         'Removes George
                              PicBack

                         'Places the new Possition and Picture (Adding George)
                              intPos = intPos + Direction
                              strGround = strCell(intPos)

                              frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                              strCell(intPos) = "91"
                              UpdateSteps
               'Either <Key on Grass> or <Key on Cement>
                    ElseIf CInt(strCell(intPos + Direction)) = 2 Or CInt(strCell(intPos + Direction)) = 3 Then
          
                         If frmMain.mnuSound.Checked = True Then
                              'Makes a picking up key sound
                                   PlaySound 0, App.Path & "\GotKey.wav"
                         End If

                         'Changes the Variable (You have one more Key)
                              intKeysNum = intKeysNum + 1
                              frmMain.lblKeysNum.Caption = intKeysNum
                              frmMain.tmrKey.Enabled = True
          
                         'Removes George (From Original Spot)
                              PicBack
          
                         'Places the new Possition and Picture (Adding George)
                              intPos = intPos + Direction
                              strGround = strCell(intPos)
                              strGround = CInt(strGround) - 2
                              strGround = "00" & strGround
          
                              frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                            strCell(intPos) = "91"
          
                              'Adds 5 points to Score
                                   dblScore = dblScore + 5
                                   frmMain.tmrPts.Enabled = False
                                   frmMain.lblPts.Caption = "5"
                                   'Repositions the Points indicator and shows it
                                        frmMain.lblPts.Top = frmMain.imgMap(intPos).Top
                                        frmMain.lblPts.Left = frmMain.imgMap(intPos).Left
                                        frmMain.lblPts.Visible = True
                                   frmMain.tmrPts.Enabled = True
                                   ScoreUpdate
                                   UpdateSteps
                                If frmMain.mnuItemInfo.Checked = True Then
                                    frmMain.lblItemInfo.Caption = "Can be used on Locked Doors and Locked Blocks"
                                    frmMain.lblItemInfo.Visible = True
                                    frmMain.tmrHideItemInfo.Enabled = False
                                    frmMain.tmrHideItemInfo.Enabled = True
                                End If
          
               '<Tile on Grass> or <Tile on Cement>
                    ElseIf CInt(strCell(intPos + Direction)) = 4 Or CInt(strCell(intPos + Direction)) = 5 Then

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a stepped on tile sound
                                   PlaySound 0, App.Path & "\Tile.wav"
                         End If
          
                         'Removes George (From Original Spot)
                              PicBack
          
                         'Places the new Possition and Picture (Adding George)
                              intPos = intPos + Direction
                              strGround = strCell(intPos)
          
                              frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                            strCell(intPos) = "091"

                                If frmMain.mnuItemInfo.Checked = True Then
                                    frmMain.lblItemInfo.Caption = "Reverses all Toggle Blocks (UP=Down, DOWN=Up)"
                                    frmMain.lblItemInfo.Visible = True
                                    frmMain.tmrHideItemInfo.Enabled = False
                                    frmMain.tmrHideItemInfo.Enabled = True
                                End If

                         'If there is a live Drone Mouse on the map
                              If booDroneMouse = True Then

                                   'If a Mouse is on a <<OFF>> Toggle Block
                                        If intDroneGround = 33 Then

                                             'Disables the AITimer
                                                  frmMain.tmrDroneAI.Enabled = False

                                             'Back to the Below Picture (Removing Mouse)
                                                  frmMain.imgMap(intDronePos).Picture = frmMain.pic(33).Picture

                                             'Sets the Grid Container accordingly
                                                  strCell(intDronePos) = "033"

                                             'The Drone Mouse Explodes
                                                  Explosion intDronePos

                                             'There is no Drone Mouse anymore
                                                  booDroneMouse = False
                                                  intDronePos = -1

                                                If frmMain.mnuItemInfo.Checked = True Then
                                                    frmMain.lblItemInfo.Caption = "You crushed him!"
                                                    frmMain.lblItemInfo.Visible = True
                                                    frmMain.tmrHideItemInfo.Enabled = False
                                                    frmMain.tmrHideItemInfo.Enabled = True
                                                End If
                                        End If

                              End If

                         'Norman is on the Map
                              If intNormanPosition >= 0 Then

                                   'If Norman is on a <<OFF>> Toggle Block
                                        If strNormanGround = "033" Then

                                             'Changes Norman's Picture
                                                  frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(895).Picture

                                             'Defeat
                                                    If frmMain.mnuItemInfo.Checked = True Then
                                                        frmMain.lblItemInfo.Caption = "Watch where your buddies are before stepping on Tiles!"
                                                        frmMain.lblItemInfo.Visible = True
                                                        frmMain.tmrHideItemInfo.Enabled = False
                                                        frmMain.tmrHideItemInfo.Enabled = True
                                                    End If
                                                  Defeat

                                        End If

                              End If
          
                              p = 1
                              For i = 0 To 399
          
                                   'Found a Toggle Block <<ON>>
                                        If CInt(strCell(i)) = 34 Then
                                             'Reverses the Blocks (If On then Off)
                                                  strCell(i) = "033"
                                                  frmMain.imgMap(i).Picture = frmMain.pic(33).Picture
                                   'Found a Toggle Block <<OFF>>
                                        ElseIf CInt(strCell(i)) = 33 Then
                                             'Reverses the Blocks (If Off then On)
                                                  strCell(i) = "034"
                                                  frmMain.imgMap(i).Picture = frmMain.pic(34).Picture
                                        End If
          
                         Next
                              UpdateSteps

               'Either <Spikes on Grass> or <Spikes on Cement>
                    ElseIf CInt(strCell(intPos + Direction)) = 6 Or CInt(strCell(intPos + Direction)) = 7 Then

                         'Removes George (From Original Spot)
                              PicBack
          
                         'Places the new Possition and Picture (Adding George)
                              intPos = intPos + Direction
                              strGround = strCell(intPos)

                              frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                            strCell(intPos) = "91"
          
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
                              If frmMain.imgBoots.Visible = False And frmMain.lblCheater.Visible = False Then
          
                                    If frmMain.mnuSound.Checked = True Then
                                         'Makes a defeat sound
                                              PlaySound 0, App.Path & "\Defeat.wav"
                                    End If
                
                                        Defeat
                                        If frmMain.mnuItemInfo.Checked = True Then
                                            frmMain.lblItemInfo.Caption = "You can only walk on Spikes with Metallic Boots!"
                                            frmMain.lblItemInfo.Visible = True
                                            frmMain.tmrHideItemInfo.Enabled = False
                                            frmMain.tmrHideItemInfo.Enabled = True
                                        End If
                              End If
                              UpdateSteps
               
               '<Locked Block>
                    ElseIf CInt(strCell(intPos + Direction)) = 8 Then

                         If intKeysNum > 0 Then
          
                                 If frmMain.mnuSound.Checked = True Then
                                      'Makes a locked block sound
                                           PlaySound 0, App.Path & "\Unlock.wav"
                                 End If
        
                                      'Changes the Variable (One less Key)
                                           intKeysNum = intKeysNum - 1
                                           frmMain.lblKeysNum.Caption = intKeysNum
                                           frmMain.tmrKey.Enabled = True
                  
                                      'Removes George
                                           PicBack
                  
                                      'Places the new Possition and Picture (Adding George)
                                           intPos = intPos + Direction
                                           strGround = strDefaultGround
        
                                           frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                                    strCell(intPos) = "91"
                         Else
                                If frmMain.mnuItemInfo.Checked = True Then
                                    frmMain.lblItemInfo.Caption = "You need a key to get past Locked Blocks"
                                    frmMain.lblItemInfo.Visible = True
                                    frmMain.tmrHideItemInfo.Enabled = False
                                    frmMain.tmrHideItemInfo.Enabled = True
                                End If
                         End If
                        UpdateSteps

               '<Water>
                    ElseIf CInt(strCell(intPos + Direction)) = 13 Then
          
                         'Removes George
                              PicBack
          
                         'Places the new Possition and Picture (Adding George)
                              intPos = intPos + Direction
                              strGround = "001"

                              frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                            strCell(intPos) = "91"

          
                         'If you have a Cement Bag
                              If intCementBagsNum > 0 Then
          
                                   If frmMain.mnuSound.Checked = True Then
                                        'Makes a stepped on water sound
                                             PlaySound 0, App.Path & "\Water.wav"
                                   End If
          
                                   'Changes the Variable (One less Cement Bag)
                                        intCementBagsNum = intCementBagsNum - 1
                                        frmMain.lblCementBagsNum.Caption = intCementBagsNum
                                        frmMain.tmrCement.Enabled = True
          
                         'If you dont have a Cement Bag
                              Else

                                    If frmMain.lblCheater.Visible = False Then
                                        If frmMain.mnuSound.Checked = True Then PlaySound 0, App.Path & "\Defeat.wav"
                                        Defeat
                                        If frmMain.mnuItemInfo.Checked = True Then
                                            frmMain.lblItemInfo.Caption = "You need Bags of Cement to get past Water!"
                                            frmMain.lblItemInfo.Visible = True
                                            frmMain.tmrHideItemInfo.Enabled = False
                                            frmMain.tmrHideItemInfo.Enabled = True
                                        End If
                                    End If
                              End If
                        UpdateSteps

               '<Locked Door>
                    ElseIf CInt(strCell(intPos + Direction)) = 23 Then
          
                         If intKeysNum > 0 Then
          
                              'Changes the Variable (One less Key)
                                   intKeysNum = intKeysNum - 1
                                   frmMain.lblKeysNum.Caption = intKeysNum
                                   frmMain.tmrKey.Enabled = True
          
                              'Removes George
                                   PicBack
          
                                   WinLevel
                        Else
                            If frmMain.mnuItemInfo.Checked = True Then
                                frmMain.lblItemInfo.Caption = "You need a Key to enter the Locked Door"
                                frmMain.lblItemInfo.Visible = True
                                frmMain.tmrHideItemInfo.Enabled = False
                                frmMain.tmrHideItemInfo.Enabled = True
                            End If
                         End If
                        UpdateSteps
          
                '<Metallic Boots>
                    ElseIf CInt(strCell(intPos + Direction)) = 35 Then
          
                         'Ensures that you don't already have the Boots
                              If frmMain.imgBoots.Visible = False Or frmMain.lblCheater.Visible = True Then
          
                                 If frmMain.mnuSound.Checked = True Then
                                      'Makes a picking up boots sound
                                           PlaySound 0, App.Path & "\GotBoots.wav"
                                 End If
                  
                                           'Shows the Boots Picture in the Items menu
                                                 frmMain.imgBoots.Visible = True
                            
                                             'Removes George
                                                 PicBack
                                 
                                            'Places the new Possition and Picture (Adding George)
                                                intPos = intPos + Direction
                                                strGround = strDefaultGround
        
                                                frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                                                strCell(intPos) = "91"
        
                                                dblScore = dblScore + 5
                                                frmMain.tmrPts.Enabled = False
                                                frmMain.lblPts.Caption = "5"
                                                'Repositions the Points indicator and shows it
                                                     frmMain.lblPts.Top = frmMain.imgMap(intPos).Top
                                                     frmMain.lblPts.Left = frmMain.imgMap(intPos).Left
                                                     frmMain.lblPts.Visible = True
                                                frmMain.tmrPts.Enabled = True
                                                ScoreUpdate
                                                If frmMain.mnuItemInfo.Checked = True And frmMain.lblCheater.Visible = False Then
                                                    frmMain.lblItemInfo.Caption = "You can now freely walk on Spikes"
                                                    frmMain.lblItemInfo.Visible = True
                                                    frmMain.tmrHideItemInfo.Enabled = False
                                                    frmMain.tmrHideItemInfo.Enabled = True
                                                End If
                                      End If
                        UpdateSteps

                '<Clock>
                      ElseIf CInt(strCell(intPos + Direction)) = 36 Then
          
                         'Ensures that you don't already have the Clock
                              If frmMain.imgClock.Visible = False Or frmMain.lblCheater.Visible = True Then
          
                         If frmMain.mnuSound.Checked = True Then
                              'Makes a picking up clock sound
                                   PlaySound 0, App.Path & "\GotClock.wav"
                         End If
                    
                                    'Resets the Timer Clock
                                        frmMain.lblTimer.Caption = "150"
                                        frmMain.tmrWatchTime.Interval = 100
                                        frmMain.tmrWatchTime.Enabled = True
                    
                                    'Shows the Clock Picture in Items menu
                                        frmMain.imgClock.Visible = True
                                      
                                    'Removes George
                                        PicBack
                    
                                    'Places the new Possition and Picture (Adding George)
                                        intPos = intPos + Direction
                                        strGround = strDefaultGround

                                        frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                                        strCell(intPos) = "91"

                                        'Adds 5 points to Score
                                             dblScore = dblScore + 5
                                             frmMain.tmrPts.Enabled = False
                                             frmMain.lblPts.Caption = "5"
                                             'Repositions the Points indicator and shows it
                                                  frmMain.lblPts.Top = frmMain.imgMap(intPos).Top
                                                  frmMain.lblPts.Left = frmMain.imgMap(intPos).Left
                                                  frmMain.lblPts.Visible = True
                                             frmMain.tmrPts.Enabled = True
                                             ScoreUpdate
                                        If frmMain.mnuItemInfo.Checked = True Then
                                            frmMain.lblItemInfo.Caption = "Time remaining reset to 150"
                                            frmMain.lblItemInfo.Visible = True
                                            frmMain.tmrHideItemInfo.Enabled = False
                                            frmMain.tmrHideItemInfo.Enabled = True
                                        End If
                              End If
                        UpdateSteps

                '<Bomb>
                      ElseIf CInt(strCell(intPos + Direction)) = 37 Then
                         
                         'Makes sure that you don't already have a Bomb
                              If frmMain.imgBomb.Visible = False Or frmMain.lblCheater.Visible = True Then
                       
                                 If frmMain.mnuSound.Checked = True Then
                                      'Makes a picking up bomb sound
                                           PlaySound 0, App.Path & "\GotBomb.wav"
                                 End If
                  
                                           'Shows the Bomb Picture in Items menu
                                                frmMain.imgBomb.Visible = True
                  
                                           'Removes George
                                                PicBack
                  
                                           'Places the new Possition and Picture (Adding George)
                                                intPos = intPos + Direction
                                                strGround = strDefaultGround
        
                                                frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                                                strCell(intPos) = "91"
        
                                           'Adds 5 points to Score
                                                dblScore = dblScore + 5
                                                frmMain.tmrPts.Enabled = False
                                                frmMain.lblPts.Caption = "5"
                                                'Repositions the Points indicator and shows it
                                                     frmMain.lblPts.Top = frmMain.imgMap(intPos).Top
                                                     frmMain.lblPts.Left = frmMain.imgMap(intPos).Left
                                                     frmMain.lblPts.Visible = True
                                                frmMain.tmrPts.Enabled = True
                                                ScoreUpdate
                                                If frmMain.mnuItemInfo.Checked = True Then
                                                    frmMain.lblItemInfo.Caption = "To use a Bomb, press <Spacebar>"
                                                    frmMain.lblItemInfo.Visible = True
                                                    frmMain.tmrHideItemInfo.Enabled = False
                                                    frmMain.tmrHideItemInfo.Enabled = True
                                                End If
                                Else
                                        If frmMain.mnuItemInfo.Checked = True Then
                                            frmMain.lblItemInfo.Caption = "You can only carry 1 Bomb at a time"
                                            frmMain.lblItemInfo.Visible = True
                                            frmMain.tmrHideItemInfo.Enabled = False
                                            frmMain.tmrHideItemInfo.Enabled = True
                                        End If
                                End If
                        UpdateSteps

               '<Cement Bag>
                    ElseIf CInt(strCell(intPos + Direction)) = 38 Then
          
                         If frmMain.mnuSound.Checked = True Then
                              'Makes a picking up cement bag sound
                                   PlaySound 0, App.Path & "\GotClock.wav"
                         End If
          
                         'Changes the Variable (You have one more Cement Bag)
                              intCementBagsNum = intCementBagsNum + 1
                              frmMain.lblCementBagsNum.Caption = intCementBagsNum
                              frmMain.tmrCement.Enabled = True
          
                         'Removes George
                              PicBack
          
                         'Places the new Possition and Picture (Adding George)
                              intPos = intPos + Direction
                              strGround = strDefaultGround

                                frmMain.imgMap(intPos).Picture = frmMain.pic(Picture).Picture
                                strCell(intPos) = "91"

                         'Adds 5 points to Score
                              dblScore = dblScore + 5
                              frmMain.tmrPts.Enabled = False
                              frmMain.lblPts.Caption = "5"
                              'Repositions the Points indicator and shows it
                                   frmMain.lblPts.Top = frmMain.imgMap(intPos).Top
                                   frmMain.lblPts.Left = frmMain.imgMap(intPos).Left
                                   frmMain.lblPts.Visible = True
                              frmMain.tmrPts.Enabled = True
                              ScoreUpdate
                        UpdateSteps
                        If frmMain.mnuItemInfo.Checked = True Then
                            frmMain.lblItemInfo.Caption = "Each Bag of Cement will allow you to walk on Water"
                            frmMain.lblItemInfo.Visible = True
                            frmMain.tmrHideItemInfo.Enabled = False
                            frmMain.tmrHideItemInfo.Enabled = True
                        End If

                  '<Drone Mouse or Death Mouse>
                         ElseIf CInt(strCell(intPos + Direction)) = 991 Or CInt(strCell(intPos + Direction)) = 992 Then

                              'Removes George (From Original Spot)
                                   PicBack

                              'Places the new Possition and Picture (Adding George)
                                   intPos = intPos + Direction
                                   strGround = strCell(intPos)

                                   frmMain.imgMap(intPos).Picture = frmMain.pic(95).Picture
                                   strCell(intPos) = "95"

                              If frmMain.mnuSound.Checked = True Then
                                   'Makes a defeat sound
                                        PlaySound 0, App.Path & "\Defeat.wav"
                              End If

                                   Defeat
                             UpdateSteps
                            If frmMain.mnuItemInfo.Checked = True Then
                                frmMain.lblItemInfo.Caption = "Watch out for Mice!"
                                frmMain.lblItemInfo.Visible = True
                                frmMain.tmrHideItemInfo.Enabled = False
                                frmMain.tmrHideItemInfo.Enabled = True
                            End If
                         End If

'_________________________________________________________________________________________

     'Norman is moving
          Else

               'Either <Grass> <Cement> <Toggle Block OFF> or <Spikes> (Grass or Cement)
                    If CInt(strCell(intNormanPosition + Direction)) < 2 Or _
                    CInt(strCell(intNormanPosition + Direction)) = 33 Or _
                    CInt(strCell(intNormanPosition + Direction)) = 6 Or _
                    CInt(strCell(intNormanPosition + Direction)) = 7 Then

                         'Back to the Below Picture (Removing Norman)
                              frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(CInt(strNormanGround)).Picture

                         'Sets the Grid Container accordingly
                              strCell(intNormanPosition) = strNormanGround

                              'Places the new Possition and Picture (Adding Norman)
                                   intNormanPosition = intNormanPosition + Direction
                                   strNormanGround = strCell(intNormanPosition)
                                   frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(Picture).Picture
                                   strCell(intNormanPosition) = "891"
                                ' If on spikes ...
                                If CInt(strCell(intNormanPosition + Direction)) = 7 Or CInt(strCell(intNormanPosition + Direction)) = 6 Then
                                    If frmMain.mnuSound.Checked = True Then
                                         'Makes a stepped on spikes sound
                                              r = RndBetween(1, 2)
                                              If r < 1.5 Then
                                                   PlaySound 0, App.Path & "\Spikes1.wav"
                                              Else
                                                   PlaySound 0, App.Path & "\Spikes2.wav"
                                              End If
                                    End If
                                End If

               '<Tile on Grass> or <Tile on Cement>
                    ElseIf CInt(strCell(intNormanPosition + Direction)) = 4 Or CInt(strCell(intNormanPosition + Direction)) = 5 Then

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a stepped on tile sound
                                   PlaySound 0, App.Path & "\Tile.wav"
                         End If

                         'Back to the Below Picture (Removing Norman)
                              frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(CInt(strNormanGround)).Picture

                         'Sets the Grid Container accordingly
                              strCell(intNormanPosition) = strNormanGround

                              'Places the new Possition and Picture (Adding Norman)
                                   intNormanPosition = intNormanPosition + Direction
                                   strNormanGround = strCell(intNormanPosition)
                                   frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(Picture).Picture
                                   strCell(intNormanPosition) = "891"

                                If frmMain.mnuItemInfo.Checked = True Then
                                    frmMain.lblItemInfo.Caption = "Reverses all Toggle Blocks (UP=Down, DOWN=Up)"
                                    frmMain.lblItemInfo.Visible = True
                                    frmMain.tmrHideItemInfo.Enabled = False
                                    frmMain.tmrHideItemInfo.Enabled = True
                                End If

                         'If there is a live Drone Mouse on the map
                              If booDroneMouse = True Then

                                   'If a Mouse is on a <<OFF>> Toggle Block
                                        If intDroneGround = 33 Then
          
                                             'Disables the AITimer
                                                  frmMain.tmrDroneAI.Enabled = False
          
                                             'Back to the Below Picture (Removing Mouse)
                                                  frmMain.imgMap(intDronePos).Picture = frmMain.pic(33).Picture
          
                                             'Sets the Grid Container accordingly
                                                  strCell(intDronePos) = "033"

                                             'The Drone Mouse Explodes
                                                  Explosion intDronePos

                                             'There is no Drone Mouse anymore
                                                  booDroneMouse = False
                                                  intDronePos = -1
          
                                                If frmMain.mnuItemInfo.Checked = True Then
                                                    frmMain.lblItemInfo.Caption = "You crushed him!"
                                                    frmMain.lblItemInfo.Visible = True
                                                    frmMain.tmrHideItemInfo.Enabled = False
                                                    frmMain.tmrHideItemInfo.Enabled = True
                                                End If
          
                                        End If
          
                              End If

                              'If George is on a <<OFF>> Toggle Block
                                   If strGround = "033" Then

                                        'Changes Norman's Picture
                                             frmMain.imgMap(intNormanPosition).Picture = frmMain.pic(895).Picture

                                        'Defeat
                                            Defeat
                                            If frmMain.mnuItemInfo.Checked = True Then
                                                frmMain.lblItemInfo.Caption = "Watch where your buddies are before stepping on Tiles!"
                                                frmMain.lblItemInfo.Visible = True
                                                frmMain.tmrHideItemInfo.Enabled = False
                                                frmMain.tmrHideItemInfo.Enabled = True
                                            End If
                                   End If

                              p = 1
                              For i = 0 To 399
          
                                   'Found a Toggle Block <<ON>>
                                        If CInt(strCell(i)) = 34 Then
                                             'Reverses the Blocks (If On then Off)
                                                  strCell(i) = "033"
                                                  frmMain.imgMap(i).Picture = frmMain.pic(33).Picture
                                   'Found a Toggle Block <<OFF>>
                                        ElseIf CInt(strCell(i)) = 33 Then
                                             'Reverses the Blocks (If Off then On)
                                                  strCell(i) = "034"
                                                  frmMain.imgMap(i).Picture = frmMain.pic(34).Picture
                                        End If

                         Next

                    End If

          End If
End Sub


Public Sub PicBack()

     'Back to the Below Picture (Removing George)
          frmMain.imgMap(intPos).Picture = frmMain.pic(CInt(strGround)).Picture

     'Sets the Grid Container accordingly
          strCell(intPos) = strGround

End Sub


'  Checks to see if you've made a Best Time
Public Sub TopTimesChecker()
  Dim i As Integer

     'If you're using Cheats, leave
          If frmSplash.lblCheat.Visible = True Then Exit Sub

     For i = 0 To 4
          If CInt(frmMain.lblTimer.Caption) > intTopTimes(i) Then
               frmMain.lblCongratulations.Caption = "Congratulations!  Enter your name into the High score archives!"
               frmMain.cmdAction.Caption = "Update"
               frmMain.txtCurrentScore.Top = frmMain.lblTopName(i).Top
               frmMain.txtCurrentScore.Visible = True
               frmMain.lblTopScore(i).BackColor = &HC0&
               frmMain.lblTopScore(i).Caption = frmMain.lblTimer

               'First place
                    If i = 0 Then
                         frmMain.lblTopName(4).Caption = frmMain.lblTopName(3).Caption
                         frmMain.lblTopName(3).Caption = frmMain.lblTopName(2).Caption
                         frmMain.lblTopName(2).Caption = frmMain.lblTopName(1).Caption
                         frmMain.lblTopName(1).Caption = frmMain.lblTopName(0).Caption

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a "High Score" Sound
                                   PlaySound 0, App.Path & "\HighScore.wav"
                         End If

                    'Stops any Music
                         frmSplash.medMidi.URL = ""
                    
                    'Disables the menus
                         frmMain.mnuFile.Enabled = False
                         frmMain.mnuOptions.Enabled = False
                         frmMain.mnuHelp.Enabled = False

               'Second place
                    ElseIf i = 1 Then
                         frmMain.lblTopName(4).Caption = frmMain.lblTopName(3).Caption
                         frmMain.lblTopName(3).Caption = frmMain.lblTopName(2).Caption
                         frmMain.lblTopName(2).Caption = frmMain.lblTopName(1).Caption

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a "High Score Sound"
                                   PlaySound 0, App.Path & "\HighScore.wav"
                         End If

                    'Stops any Music
                         frmSplash.medMidi.URL = ""

                    'Disables the menus
                         frmMain.mnuFile.Enabled = False
                         frmMain.mnuOptions.Enabled = False
                         frmMain.mnuHelp.Enabled = False

               'Third place
                    ElseIf i = 2 Then
                         frmMain.lblTopName(4).Caption = frmMain.lblTopName(3).Caption
                         frmMain.lblTopName(3).Caption = frmMain.lblTopName(2).Caption
                         frmMain.lblTopName(2).Caption = frmMain.lblTopName(1).Caption

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a "High Score Sound"
                                   PlaySound 0, App.Path & "\HighScore.wav"
                         End If

                    'Stops any Music
                         frmSplash.medMidi.URL = ""

                    'Disables the menus
                         frmMain.mnuFile.Enabled = False
                         frmMain.mnuOptions.Enabled = False
                         frmMain.mnuHelp.Enabled = False

               'Fourth place
                    ElseIf i = 3 Then
                         frmMain.lblTopName(4).Caption = frmMain.lblTopName(3).Caption

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a "High Score Sound"
                                   PlaySound 0, App.Path & "\HighScore.wav"
                         End If

                    'Stops any Music
                         frmSplash.medMidi.URL = ""

                    'Disables the menus
                         frmMain.mnuFile.Enabled = False
                         frmMain.mnuOptions.Enabled = False
                         frmMain.mnuHelp.Enabled = False

                    End If

               frmMain.fraHighScore.Visible = True
               frmMain.txtCurrentScore.Text = strBTName
               frmMain.txtCurrentScore.SelStart = 0
               frmMain.txtCurrentScore.SelLength = Len(frmMain.txtCurrentScore.Text)
               intTopTimePlace = i
               Exit Sub

          Else
     
                    If frmMain.mnuSound.Checked = True Then PlaySound 0, App.Path & "\Victory.wav"

                    'Stops any Music
                         frmSplash.medMidi.URL = ""

          End If
     
     Next

End Sub



'  Code on Win
Public Sub WinLevel()
  Dim BlocksNum As Integer
  Dim p As Integer
  Dim i As Integer

    'Stops the Timer
        frmMain.tmrTimer.Enabled = False
    'Stops the Drone Mouse Timer
        frmMain.tmrDroneAI.Enabled = False
    'Stops the Death Mouse Timer
        frmMain.tmrDeathAI.Enabled = False
    'Checks to see if you've made a Best Time
        If blnSecret = False Then
            TopTimesChecker
        Else
            If frmMain.mnuSound.Checked = True Then PlaySound 0, App.Path & "\Victory.wav"
            frmSplash.medMidi.URL = ""
        End If
    'If it is the Last Level
        If strNextLevel = "Last" And frmMain.fraHighScore.Visible = False Then
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
    'If it isn't the Last Level
        Else
            If frmMain.fraHighScore.Visible = False Then
                'The name of the file to load:
                    strFile = strNextLevel
                    If Mid$(strFile, Len(strFile) - 3, 4) = ".lvl" Then strFile = App.Path & "\Scenarios\" & strNextLevel
                    If Mid$(strFile, Len(strFile) - 3, 4) = ".cus" Then strFile = strBestTimesDir & "\" & strNextLevel
                    
                'Sets the Score
                    dblScore = dblScore + CDbl(frmMain.lblTimer.Caption) + (CDbl(frmMain.lblTimer.Caption) \ 2)
                    ScoreUpdate
                'Resets the Lives
                    If frmSplash.lblCheat.Visible = False Then intLivesNum = 5
                    If frmSplash.lblCheat.Visible = True Then intLivesNum = 10
                'Resets the Menu Properties acordingly
                    frmMain.mnuQuitGame.Enabled = True
                    frmMain.mnuSaveGame.Enabled = False
                    frmMain.mnuSaveGameAs.Enabled = False
                    frmMain.mnuPauseGame.Enabled = False
                    frmMain.mnuRestartLevel.Enabled = False
                    frmMain.mnuBestTimes.Enabled = True
                'A Game not in progress
                    booPlaying = False
                'The game is not paused
                    frmMain.fraPaused.Visible = False
                'Loads the Level
                    LoadLevel
            End If
        End If
    'Playing secret level
        If blnSecret = True Then
            'No cheating detected ...
                If blnValidSecret = True Then
                    If frmMain.lblTimer.Caption = "110" Then
                        BlocksNum = 0
                        p = 1
                        For i = 0 To 399
                            If CInt(strCell(i)) = 9 Then
                                BlocksNum = BlocksNum + 1
                            End If
                            p = p + 3
                            frmCreate.imgMap(i).Enabled = False
                        Next
                        If BlocksNum = 4 Then
                            frmMain.lblSecret(0).Visible = True
                            frmMain.lblSecret(1).Visible = True
                            frmMain.lblSecret(2).Visible = True
                        End If
                    End If
            'Cheating detected!
                Else
                    MsgBox "You think we don't know when you cheat?!" & vbNewLine & "It was a good try though.", vbCritical, "Marque Castle"
                End If
        End If
End Sub


'Makes an Explosion
Public Sub Explosion(PlayerPos As Integer)
  With frmMain

     'Places the Explosion Picture's Position
          .imgExplosion.Left = .imgMap(PlayerPos).Left - 240
          .imgExplosion.Top = .imgMap(PlayerPos).Top - 240

          .imgExplosionSmall.Left = .imgExplosion.Left + 255
          .imgExplosionSmall.Top = .imgExplosion.Top + 255

     'Shows the small Explosion Picture
         .imgExplosionSmall.Visible = True
         .tmrExplosionSmall.Enabled = False: .tmrExplosionSmall.Enabled = True


     'Destroys the Blocks within the radius
          If PlayerPos - 21 >= 0 Then
          
               'Found a Block
                    If strCell(PlayerPos - 21) = "009" Then
                         strCell(PlayerPos - 21) = strDefaultGround
                         .imgMap(PlayerPos - 21).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 25
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "25"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate

               'Found a Mouse
                    ElseIf strCell(PlayerPos - 21) = "991" Or strCell(PlayerPos - 21) = "992" Or strCell(PlayerPos - 21) = "993" Then
                        If strCell(PlayerPos - 21) = "991" Then
                            .tmrDroneAI.Enabled = False
                            booDroneMouse = False
                            intDronePos = -1
                        Else
                            .tmrDeathAI.Enabled = False
                            blnDeathMouse = False
                            intDeathPos = -1
                        End If
                        strCell(PlayerPos - 21) = strDefaultGround
                        .imgMap(PlayerPos - 21).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 150
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "150"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate
                    End If
          End If

          If PlayerPos - 20 >= 0 Then

               'Found a Block
                    If strCell(PlayerPos - 20) = "009" Then
                         strCell(PlayerPos - 20) = strDefaultGround
                         .imgMap(PlayerPos - 20).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 15
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "15"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate

               'Found a Mouse
                    ElseIf strCell(PlayerPos - 20) = "991" Or strCell(PlayerPos - 20) = "992" Or strCell(PlayerPos - 20) = "993" Then
                        If strCell(PlayerPos - 201) = "991" Then
                            .tmrDroneAI.Enabled = False
                            booDroneMouse = False
                            intDronePos = -1
                        Else
                            .tmrDeathAI.Enabled = False
                            blnDeathMouse = False
                            intDeathPos = -1
                        End If
                        strCell(PlayerPos - 20) = strDefaultGround
                        .imgMap(PlayerPos - 20).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 150
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "150"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate
                    End If
          End If

          If PlayerPos - 19 >= 0 Then

               'Found a Block
                    If strCell(PlayerPos - 19) = "009" Then
                         strCell(PlayerPos - 19) = strDefaultGround
                         .imgMap(PlayerPos - 19).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 15
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "15"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate

               'Found a Mouse
                    ElseIf strCell(PlayerPos - 19) = "991" Or strCell(PlayerPos - 19) = "992" Or strCell(PlayerPos - 19) = "993" Then
                        If strCell(PlayerPos - 19) = "991" Then
                            .tmrDroneAI.Enabled = False
                            booDroneMouse = False
                            intDronePos = -1
                        Else
                            .tmrDeathAI.Enabled = False
                            blnDeathMouse = False
                            intDeathPos = -1
                        End If
                        strCell(PlayerPos - 19) = strDefaultGround
                        .imgMap(PlayerPos - 19).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 100
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "100"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate
                    End If
          End If

          If PlayerPos - 1 >= 0 Then

               'Found a Block
                    If strCell(PlayerPos - 1) = "009" Then
                         strCell(PlayerPos - 1) = strDefaultGround
                         .imgMap(PlayerPos - 1).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 15
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "15"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate

               'Found a Mouse
                    ElseIf strCell(PlayerPos - 1) = "991" Or strCell(PlayerPos - 1) = "992" Or strCell(PlayerPos - 1) = "993" Then
                        If strCell(PlayerPos - 1) = "991" Then
                            .tmrDroneAI.Enabled = False
                            booDroneMouse = False
                            intDronePos = -1
                        Else
                            .tmrDeathAI.Enabled = False
                            blnDeathMouse = False
                            intDeathPos = -1
                        End If
                        strCell(PlayerPos - 1) = strDefaultGround
                        .imgMap(PlayerPos - 1).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 150
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "150"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate
                    End If
          End If

          If PlayerPos + 1 <= 399 Then

               'Found a Block
                    If strCell(PlayerPos + 1) = "009" Then
                         strCell(PlayerPos + 1) = strDefaultGround
                         .imgMap(PlayerPos + 1).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 15
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "15"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate

               'Found a Mouse
                    ElseIf strCell(PlayerPos + 1) = "991" Or strCell(PlayerPos + 1) = "992" Or strCell(PlayerPos + 1) = "993" Then
                        If strCell(PlayerPos + 1) = "991" Then
                            .tmrDroneAI.Enabled = False
                            booDroneMouse = False
                            intDronePos = -1
                        Else
                            .tmrDeathAI.Enabled = False
                            blnDeathMouse = False
                            intDeathPos = -1
                        End If
                        strCell(PlayerPos + 1) = strDefaultGround
                        .imgMap(PlayerPos + 1).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 150
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "150"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate
                    End If
          End If

          If PlayerPos + 19 <= 399 Then

               'Found a Block
                    If strCell(PlayerPos + 19) = "009" Then
                         strCell(PlayerPos + 19) = strDefaultGround
                         .imgMap(PlayerPos + 19).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 15
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "15"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate

               'Found a Mouse
                    ElseIf strCell(PlayerPos + 19) = "991" Or strCell(PlayerPos + 19) = "992" Or strCell(PlayerPos + 19) = "993" Then
                        If strCell(PlayerPos + 19) = "991" Then
                            .tmrDroneAI.Enabled = False
                            booDroneMouse = False
                            intDronePos = -1
                        Else
                            .tmrDeathAI.Enabled = False
                            blnDeathMouse = False
                            intDeathPos = -1
                        End If
                        strCell(PlayerPos + 19) = strDefaultGround
                        .imgMap(PlayerPos + 19).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 150
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "150"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate
                    End If
          End If

          If PlayerPos + 20 <= 399 Then

               'Found a Block
                    If strCell(PlayerPos + 20) = "009" Then
                         strCell(PlayerPos + 20) = strDefaultGround
                         .imgMap(PlayerPos + 20).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 15
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "15"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate

               'Found a Mouse
                    ElseIf strCell(PlayerPos + 20) = "991" Or strCell(PlayerPos + 20) = "992" Or strCell(PlayerPos + 20) = "993" Then
                        If strCell(PlayerPos + 20) = "991" Then
                            .tmrDroneAI.Enabled = False
                            booDroneMouse = False
                            intDronePos = -1
                        Else
                            .tmrDeathAI.Enabled = False
                            blnDeathMouse = False
                            intDeathPos = -1
                        End If
                        strCell(PlayerPos + 20) = strDefaultGround
                        .imgMap(PlayerPos + 20).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 150
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "150"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate
                    End If
          End If

          If PlayerPos + 21 <= 399 Then

               'Found a Block
                    If strCell(PlayerPos + 21) = "009" Then
                         strCell(PlayerPos + 21) = strDefaultGround
                         .imgMap(PlayerPos + 21).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 15
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "15"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate

               'Found a Mouse
                    ElseIf strCell(PlayerPos + 21) = "991" Or strCell(PlayerPos + 21) = "992" Or strCell(PlayerPos + 21) = "993" Then
                        If strCell(PlayerPos + 21) = "991" Then
                            .tmrDroneAI.Enabled = False
                            booDroneMouse = False
                            intDronePos = -1
                        Else
                            .tmrDeathAI.Enabled = False
                            blnDeathMouse = False
                            intDeathPos = -1
                        End If
                        strCell(PlayerPos + 21) = strDefaultGround
                        .imgMap(PlayerPos + 21).Picture = .pic(CInt(strDefaultGround)).Picture
                         dblScore = dblScore + 150
                         .tmrPts.Enabled = False
                         .lblPts.Caption = "150"
                         'Repositions the Points indicator and shows it
                              .lblPts.Top = .imgMap(intPos).Top
                              .lblPts.Left = .imgMap(intPos).Left
                              .lblPts.Visible = True
                         .tmrPts.Enabled = True
                         ScoreUpdate
                    End If
          End If

    'Shows the large explosion picture
    .imgExplosion.Visible = True
    .tmrExplosion.Enabled = False: .tmrExplosion.Enabled = True
  End With
End Sub


'  Updates your Score Board
Public Sub ScoreUpdate()
  Dim i As Integer

    On Error Resume Next

  'If you're using Cheats, you get an extra 40pts every time!
  If frmSplash.lblCheat.Visible = True Then dblScore = dblScore + 40
  'Adds the rest of the actual score
  dblScore = dblScore + CDbl(frmMain.lblPts.Caption)

  'Output the score
  frmMain.lblScore.Caption = Format$(dblScore, "0000000000")

  frmMain.lblScore.ForeColor = &HFF&  'Changes the colour of the score numbers
  frmMain.tmrScore.Enabled = True     'Changes the colour back
End Sub
     
Public Sub QuitGame()
  Dim i As Integer

     If frmMain.mnuSound.Checked = True Then
          'Makes a beeping sound
               PlaySound 0, App.Path & "\Beep.wav"
     End If
          
          'You're no longer playing
               booPlaying = False

          'Stops and Resets the clock
               frmMain.tmrTimer.Enabled = False
               frmMain.lblTimer.Caption = "150"

          'Hides the map
               For i = 0 To 399
                    frmMain.imgMap(i).Visible = False
               Next

          'Hides the appropriate status displays
               frmMain.fraLevelInfo.Visible = False
               frmMain.fraItems.Visible = False
               frmMain.fraTimer.Visible = False
               frmMain.fraSkinDir.Visible = False
               frmMain.fraMessage.Visible = False

          'Resets the Menu Properties acordingly
               frmMain.mnuQuitGame.Enabled = False
               frmMain.mnuSaveGame.Enabled = False
               frmMain.mnuSaveGameAs.Enabled = False
               frmMain.mnuPauseGame.Enabled = False
               frmMain.mnuRestartLevel.Enabled = False
               frmMain.mnuBestTimes.Enabled = False
               frmMain.fraPaused.Visible = False
               frmMain.fraDefeat.Visible = False

          'You no longer have any Keys
               intKeysNum = 0
               frmMain.lblKeysNum.Caption = intKeysNum
          
          'You no longer have any Cement Bags
               intCementBagsNum = 0
               frmMain.lblCementBagsNum.Caption = intCementBagsNum
          
          'Resets your score
               dblScore = 0
               ScoreUpdate

          'Stops any Music
               frmSplash.medMidi.URL = ""

End Sub


'  Error Handler
Public Function FileExists(strFile As String) As Boolean
  Dim bRepeat As Boolean

  On Error Resume Next

  Do
    Err.Clear
    Close
    bRepeat = False

    If frmMain.fraPictures.Caption = "Read" Then
      Open strFile For Input As #1
    Else
      Open strFile For Output As #1
    End If
    
    If Err.Number Then
      If MsgBox("Can't read or write file:" & vbNewLine & """" & strFile & """." & vbNewLine & vbNewLine & "Retry?", vbYesNo, "Can't open file") = vbYes Then bRepeat = True
    End If
  Loop While bRepeat
End Function

Private Sub UpdateSteps()
  'Adds the step
  dblSteps = dblSteps + 1

  'Output the number of steps
  frmMain.lblSteps(1).Caption = Format$(dblSteps + 1, "0000000000")

  'Changes the colour of the score numbers
  frmMain.lblSteps(1).ForeColor = &HFF&

  'Changes the colour back
  frmMain.tmrScore.Enabled = False
  frmMain.tmrScore.Enabled = True
End Sub
