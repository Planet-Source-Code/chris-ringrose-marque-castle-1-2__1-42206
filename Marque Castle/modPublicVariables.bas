Attribute VB_Name = "modPublicVariables"


'Marque Castle v1.2
'modPublicVariables
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


Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)

'   <<Save and Load Variables>>

    'The frame for the ending credis
          Public intCreditNum As Integer

    'The actual Level's Name
          Public strLevelFile As String

     'Determines the File location to Load
          Public strFile As String

     'The current Level
          Public strLevel As String

     'The level after the current one
         Public strNextLevel As String

     'The title of the current level
         Public strLevelTitle As String

     'The Message presented to the Gamer at Startup
         Public strMessage As String

     'The map of each picture
         Public strData As String

     'The Author of the map
         Public strAuthor As String

     'The Ground Below your feet
         Public strGround As String

    'The Ground below the Death Mouse's feet
        Public strDeathGround As String

    'The Default Ground below your feet
        Public strDefaultGround As String


'   <<Used for playing the Game>>
    'Playing the secret level or not
        Public blnSecret As Boolean
        Public blnValidSecret As Boolean

     'The best times name entered
         Public strBTName As String

     'Norman's Position
          Public intNormanPosition As Integer

     'Norman's Ground Pic
          Public strNormanGround As String

     'True if a Game is in progress
         Public booPlaying As Boolean

     'The directory from which the pictures are loaded
         Public strSkinDir As String

     'The Score
          Public dblScore As Double
          
    ' The Steps
         Public dblSteps As Double

     'The Top Times
          Public strTopNames(4) As String
          Public intTopTimes(4)  As Integer

     'The Custom Information
          Public strProperties(4)

     'George's Possion:
          Public intPos As Integer

     'Your number of Lives
          Public intLivesNum As Integer

     'Your number of Keys
          Public intKeysNum As Integer

     'Number of Cement Bags
          Public intCementBagsNum As Integer

     'There is a live Drone Mouse on the map
          Public booDroneMouse As Boolean
          
     'There is a live Death Mouse on the map
         Public blnDeathMouse As Boolean

     'The Top Time Place
          Public intTopTimePlace As Integer

     'Whether Norman's on the Map or not
          Public booNorman As Boolean

'  <<Used for creating your own Scenario>>

     'What Picture file is on each block
          Public strCell(399) As String

        Public strBestTimesDir As String
