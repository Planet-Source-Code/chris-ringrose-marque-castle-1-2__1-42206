Attribute VB_Name = "modDecrypted"


'Marque Castle v1.2
'modDecrypted
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


'  Decrypts the Information in CusSet.opt
Public Sub DecryptedCusSet(strDecryptInfo As String)
  Dim i As Integer
  
     'Replaces each symbol with
          For i = 1 To Len(strProperties(CInt(frmMain.lblStrProperties.Caption)))

               Select Case Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1)

                    Case Is = "*"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "a"

                    Case Is = "!"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "b"

                    Case Is = "~"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "c"

                    Case Is = "+"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "d"

                    Case Is = "&"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "e"

                    Case Is = "'"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "f"

                    Case Is = "="
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "g"

                    Case Is = ":"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "h"

                    Case Is = "@"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "i"

                    Case Is = "#"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "j"

                    Case Is = "^"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "k"

                    Case Is = "("
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "l"

                    Case Is = "]"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "m"

                    Case Is = ")"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "n"

                    Case Is = "["
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "o"

                    Case Is = ";"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "p"

                    Case Is = ","
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "q"

                    Case Is = ">"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "r"

                    Case Is = "?"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "s"

                    Case Is = "y"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "t"

                    Case Is = "\"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "u"

                    Case Is = "/"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "v"

                    Case Is = "G"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "w"

                    Case Is = "<"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "x"

                    Case Is = "%"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "y"

                    Case Is = "6"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "z"

                    Case Is = "D"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "\"

                    Case Is = "3"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "."

                    Case Is = "I"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = ":"

                    Case Is = "b"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "="

                    Case Is = "8"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = ","

                    Case Is = "m"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "0"

                    Case Is = "}"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "1"

                    Case Is = "-"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "2"

                    Case Is = "_"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "3"

                    Case Is = "`"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "4"

                    Case Is = "x"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "5"

                    Case Is = "$"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "6"

                    Case Is = "o"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "7"

                    Case Is = "j"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "8"

                    Case Is = "{"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "9"

                    Case Is = "V"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = " "

               End Select

          Next

End Sub


'_____________________________________________________________________________________________

'  Encrypts the Information in CusSet.opt
Public Sub EncryptedCusSet(strEcryptInfo As String)
  Dim i As Integer
  
     'Replaces each symbol with
          For i = 1 To Len(strProperties(CInt(frmMain.lblStrProperties.Caption)))

               Select Case Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1)

                    Case Is = "a"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "*"

                    Case Is = "b"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "!"

                    Case Is = "c"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "~"

                    Case Is = "d"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "+"

                    Case Is = "e"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "&"

                    Case Is = "f"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "'"

                    Case Is = "g"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "="

                    Case Is = "h"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = ":"

                    Case Is = "i"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "@"

                    Case Is = "j"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "#"

                    Case Is = "k"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "^"

                    Case Is = "l"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "("

                    Case Is = "m"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "]"

                    Case Is = "n"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = ")"

                    Case Is = "o"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "["

                    Case Is = "p"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = ";"

                    Case Is = "q"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = ","

                    Case Is = "r"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = ">"

                    Case Is = "s"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "?"

                    Case Is = "t"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "y"

                    Case Is = "u"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "\"

                    Case Is = "v"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "/"

                    Case Is = "w"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "G"

                    Case Is = "x"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "<"

                    Case Is = "y"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "%"

                    Case Is = "z"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "6"

                    Case Is = "\"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "D"

                    Case Is = "."
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "3"

                    Case Is = ":"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "I"

                    Case Is = "="
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "b"

                    Case Is = ","
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "8"

                    Case Is = "0"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "m"

                    Case Is = "1"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "}"

                    Case Is = "2"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "-"

                    Case Is = "3"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "_"

                    Case Is = "4"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "`"

                    Case Is = "5"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "x"

                    Case Is = "6"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "$"

                    Case Is = "7"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "o"

                    Case Is = "8"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "j"

                    Case Is = "9"
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "{"

                    Case Is = " "
                         Mid$(strProperties(CInt(frmMain.lblStrProperties.Caption)), i, 1) = "V"

               End Select

          Next

End Sub

