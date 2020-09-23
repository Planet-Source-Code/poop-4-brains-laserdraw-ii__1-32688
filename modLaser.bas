Attribute VB_Name = "modLaser"
Enum LDir
LFromCorner = 0 'the ways you can draw the laser
lcustom = 1
LFromRSide = 2
LFromBottom = 3
End Enum

Enum SDir
SHorizontal = 0 'the ways you can sweep
SVertical = 1
End Enum

Public Running As Boolean
Public PercentDone As Long

Sub LaserDraw(picsource As Object, picdest As Object, way As LDir, inverse As Boolean, Optional xsource As Integer, Optional ysource As Integer)
On Error Resume Next 'it might have probs with some pixels i doubt though
Dim X As Integer, Y As Integer, color As Long, sx As Integer, sy As Integer

picsource.DrawWidth = 5
picsource.Line (0, picsource.ScaleHeight)-(picsource.ScaleWidth, picsource.ScaleHeight), picdest.BackColor
picsource.Line (picsource.ScaleWidth, 0)-(picsource.ScaleWidth, picsource.ScaleHeight), picdest.BackColor

sx = picdest.ScaleWidth \ 2 'the middle
sy = picdest.ScaleHeight \ 2 'i dont know what for

If way <> LDir.LFromBottom Then
For X = 0 To picsource.ScaleWidth - 1 'trim so you dont get the white leftovers
      For Y = 0 To picsource.ScaleHeight - 1
      If Running = False Then Exit For
      
      color = picsource.Point(X, Y) 'copy the color
      
       If inverse = True Then 'if to inverse then get the inversed color
      color = InverseColor(color)
      End If
      
      Select Case way 'from where
      Case LDir.LFromRSide  'from the side
      picdest.Line (X, Y)-(picdest.ScaleWidth, Y), color
      
      Case LDir.lcustom  'the custom source coords of the laser
      picdest.Line (xsource, ysource)-(X, Y), color
         
      Case LDir.LFromCorner  'from the left-bottom corner
      picdest.Line (picdest.ScaleWidth, picdest.ScaleHeight)-(X, Y), color
      picdest.Line (picdest.ScaleWidth, picdest.ScaleHeight)-(X, Y), color
      End Select
      
      picdest.PSet (X, Y), color  'draw the pixel
      
      DoEvents
      Next Y
      
      PercentDone = (X / picsource.ScaleWidth) * 100
Next X

ElseIf way = LDir.LFromBottom Then
For Y = 0 To picsource.ScaleHeight - 1 'trim so you dont get the white leftovers
      For X = 0 To picsource.ScaleWidth - 1
      If Running = False Then Exit For
      
      color = picsource.Point(X, Y) 'copy the color
      
       If inverse = True Then 'if to inverse then get the inversed color
      color = InverseColor(color)
      End If
      
      picdest.Line (X, picdest.ScaleHeight)-(X, Y), color
      
      picdest.PSet (X, Y), color 'draw the pixel
      
      DoEvents
      Next X
      
      PercentDone = (Y / picsource.ScaleWidth) * 100
Next Y
End If
      
Running = False
frmMain.mnuStop.Enabled = False
End Sub

Sub CopyPic(pic1 As PictureBox, pic2 As PictureBox, inverse As Boolean) 'this is used to copy the picture to the preveiw
Dim X As Single, Y As Single, color As Long

For X = 0 To pic2.ScaleWidth 'everything
     For Y = 0 To pic2.ScaleWidth

     color = pic1.Point(X, Y) 'get the color
     
     If inverse = True Then color = InverseColor(color) 'if its inverse then get the inversed's color
     
     pic2.PSet (X, Y), color 'draw it on the other box
     DoEvents
     Next Y
     
Next X
End Sub

'the following two functions i got from source60.bas
Function InverseColor(OldColor) As Long
'by monk-e-god
dacolor$ = RGBtoHEX(OldColor)
redx% = Val("&H" + Right(dacolor$, 2))
greenx% = Val("&H" + Mid(dacolor$, 3, 2))
bluex% = Val("&H" + Left(dacolor$, 2))
newred% = 255 - redx%
newgreen% = 255 - greenx%
newblue% = 255 - bluex%
InverseColor = RGB(newred%, newgreen%, newblue%)
End Function
Function RGBtoHEX(RGB)
'heh, I didnt make this one...
    a$ = Hex(RGB)
    b% = Len(a$)
    If b% = 5 Then a$ = "0" & a$
    If b% = 4 Then a$ = "00" & a$
    If b% = 3 Then a$ = "000" & a$
    If b% = 2 Then a$ = "0000" & a$
    If b% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function
'*********************************

Sub SweepDraw(picsource As Object, picdest As Object, way As SDir, inverse As Boolean)
On Error Resume Next 'it might have probs with some pixels i doubt though
Dim X As Integer, Y As Integer, color As Long

Select Case way
Case SDir.SHorizontal
For X = 0 To picsource.ScaleWidth - 1 'trim so you dont get the white leftovers
      For Y = 0 To picsource.ScaleHeight - 1
      If Running = False Then Exit For
      
      color = picsource.Point(X, Y) 'copy the color
      
      If inverse = True Then 'if to inverse then get the inversed color
      color = InverseColor(color)
      End If
      
      picdest.PSet (X, Y), color 'draw the pixel
      
      DoEvents
      Next Y
    PercentDone = (X / picsource.ScaleWidth) * 100
Next X
Case SDir.SVertical
For Y = 0 To picsource.ScaleHeight - 1 'trim so you dont get the white leftovers
      For X = 0 To picsource.ScaleWidth - 1
      If Running = False Then Exit For
      
      color = picsource.Point(X, Y) 'copy the color
      
      If inverse = True Then 'if to inverse then get the inversed color
      color = InverseColor(color)
      End If
      
      picdest.PSet (X, Y), color 'draw the pixel
      
      DoEvents
      Next X
            PercentDone = (Y / picsource.ScaleHeight) * 100
Next Y
End Select

Running = False
frmMain.mnuStop.Enabled = False
End Sub
