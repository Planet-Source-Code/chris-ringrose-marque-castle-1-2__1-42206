Attribute VB_Name = "modAI"


'Marque Castle v1.2
'modAI
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



'The Drone Mouse's ground
     Public intDroneGround As Integer

'The Death Mouse's ground
     Public intDeathGround As Integer

'The Drone Mouse's Direction facing  (0-Right,  1-Up,  2-Left,  3-Down)
     Public intDroneDir As Integer

'The Drone Mouse's Position
     Public intDronePos As Integer

'The Death Mouse's Position
     Public intDeathPos As Integer

'Counts the number of times the Drone Mouse could not move
     Public intDroneUnmove As Integer

'Counts the number of times the Death Mouse could not move
     Public intDeathUnmove As Integer

'Whether the Drone Mouse has Moved
     Dim booDroneMoved As Boolean

'The block facing (0-Right, 1-up,2-left, or 3-down from Drone)
     Dim intTestDir As Integer

'The block to test (0 To 399)
     Dim intTestPos As Integer

'The Death Mouse's Last Position
     Dim intDeathLast As Integer

'The Distance (Rows) from George
     Dim intDeltaR As Integer

'The Distance (Columns) from George
     Dim intDeltaC As Integer




'  <<The AI for the Drone Mouse>>
Public Sub DroneMouseAI()
  Dim temp As Integer
  Dim p As Integer
  Dim i As Integer
  

    If frmMain.lblEnd(0).Visible = True Or frmMain.fraMessage.Visible = True Or frmMain.cmdBegin02.Visible = True Then
        frmMain.tmrDroneAI.Enabled = False
        frmMain.tmrDeathAI.Enabled = False
        Exit Sub
    End If

     On Error Resume Next
     booDroneMoved = False
     temp = -1

     'Moves the Drone Mouse to the next block
          Do
               'Makes sure it isn't on a boarder
                    Do
                         temp = temp + 1
                         intTestDir = (intDroneDir + temp + 3) Mod 4
                    Loop While (intDronePos Mod 20 = 19 And intTestDir = 0) _
                         Or (intDronePos < 20 And intTestDir = 1) _
                         Or (intDronePos Mod 20 = 0 And intTestDir = 2) _
                         Or (intDronePos > 379 And intTestDir = 3)

               If intTestDir = 0 Then
                    intTestPos = intDronePos + 1
               ElseIf intTestDir = 1 Then
                    intTestPos = intDronePos - 20
               ElseIf intTestDir = 2 Then
                    intTestPos = intDronePos - 1
               Else
                    intTestPos = intDronePos + 20
               End If
               If intTestPos < 0 Then
                    frmMain.tmrDroneAI.Enabled = False
                    Exit Sub
                End If
               'Ensures that the Drone Mouse isn't facing an obstacle
                    If CInt(strCell(intTestPos)) <= 7 Or _
                         CInt(strCell(intTestPos)) = 33 Or _
                         CInt(strCell(intTestPos)) = 91 Or _
                         CInt(strCell(intTestPos)) = 92 Or _
                         CInt(strCell(intTestPos)) = 93 Or _
                         CInt(strCell(intTestPos)) = 94 Then

                              'Moves the Drone Mouse
                              MoveMouse intDroneGround, intDronePos, 991, intTestPos

                              'Makes the Drone Mouse face the way it moved
                                   intDroneDir = intTestDir

               'The Drone Mouse is facing an obstacle
                    Else

                         intDroneDir = intDroneDir + 1
                         Exit Sub

                    End If

               'It moved into a Toggle Block (Grass or Cement)
                    If intDroneGround = 4 Or intDroneGround = 5 Then

                         If frmMain.mnuSound.Checked = True Then
                              'Makes a stepped on tile sound
                                   PlaySound 0, App.Path & "\Tile.wav"
                         End If
                         
                         'Switches the Toggle Blocks
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

                         'If George or Norman are on a Toggle Block, they die
                              If CInt(strGround) = 33 Or CInt(strNormanGround) = 33 Then

                                   'Defeat
                                      'Makes a beeping sound
                                      If frmMain.mnuSound.Checked = True Then PlaySound 0, App.Path & "\Defeat.wav"
                                      Defeat
                              End If

                    End If

               'It moved onto George
                    If intDronePos = intPos Then
     
                         'Defeat
                            'Makes a beeping sound
                            If frmMain.mnuSound.Checked = True Then PlaySound 0, App.Path & "\Defeat.wav"
                            Defeat

                            If frmMain.mnuItemInfo.Checked = True Then
                                frmMain.lblItemInfo.Caption = "Watch out for Mice!"
                                frmMain.lblItemInfo.Visible = True
                                frmMain.tmrHideItemInfo.Enabled = False
                                frmMain.tmrHideItemInfo.Enabled = True
                            End If

                    End If

               'The Drone Mouse has moved
                    booDroneMoved = True

          Loop While booDroneMoved = False And temp < 3

End Sub


'  <<Moves the Mouse>>
Public Sub MoveMouse(MouseGround As Integer, MousePos As Integer, MousePicture As Integer, MousePosTo As Integer)

    If MousePos < 0 Then
        frmMain.tmrDroneAI.Enabled = False
        Exit Sub
    End If

     'Removes Mouse (From Original Spot)
          'Back to the Below Picture (Removing Mouse)
               frmMain.imgMap(MousePos).Picture = frmMain.pic(MouseGround).Picture

          'Sets the Grid Container accordingly
               strCell(MousePos) = MouseGround

     'Places the new Possition and Picture (Adding Mouse)
          MousePos = MousePosTo
          MouseGround = strCell(MousePosTo)
          strCell(MousePos) = "991"

          frmMain.imgMap(MousePosTo).Picture = frmMain.pic(MousePicture).Picture

End Sub
