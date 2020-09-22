Attribute VB_Name = "modSoundPlayer"


'Marque Castle v1.2
'modSoundPlayer
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

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Sub PlaySound(sndEvent As Integer, Optional sndFilename As String)

     'The Sound File to Load
          Dim strSoundFileName As String

     'Applies the File Path and Name to the String
          strSoundFileName = sndFilename

     'If you don't have a Sound Card, or you have a problem with the Sound System
          On Error GoTo ErrorTrap

     'Plays the Sound
          sndPlaySound strSoundFileName, 3 'The 3 prevents the system from freezing during playback

     'The Error Trapper
ErrorTrap:

          'Nothing happens

End Sub
