<div align="center">

## A 3D cube with rotation and zoom in/out\!


</div>

### Description

This program rotates a 3d cube to the 4 directions, using a translation code,

and also has a zoom in/out option (control it with: W, A, D, X, 1 & 2)
 
### More Info
 
You have to create a Timer and left it named Timer1.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Omer Kornitz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/omer-kornitz.md)
**Level**          |Unknown
**User Rating**    |3.3 (10 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/omer-kornitz-a-3d-cube-with-rotation-and-zoom-in-out__1-3845/archive/master.zip)





### Source Code

```
Dim ww As Integer
Dim Ixy_angle, Iz_angle, dYYshift, dXXshift, csx, csy As Integer
Dim cosa, cosb, sina, sinb, coscosba, cossinba, sincosba, sinsinba, zoom, pi180 As Double
'This is the translation function
Private Sub posxy(x1 As Double, y1 As Double, z1 As Double)
    Dim Yy, Xx As Double
    Yy = zoom / (10# - (z1 * cosb + y1 * sinsinba - x1 * sincosba))
    Xx = 100# * (1# + (y1 * cosa + x1 * sina) * Yy)
    csx = Int(dXXshift) + Int(Xx)
    Xx = 100# * (1# + (y1 * cossinba - x1 * coscosba - z1 * sinb) * Yy)
    csy = Int(dYYshift) + Int(Xx)
End Sub
Sub rollup()
     Iz_angle = (Iz_angle + 5)
     cosb = Cos(Iz_angle * pi180)
     sinb = Sin(Iz_angle * pi180)
     sinsinba = sinb * sina
     sincosba = sinb * cosa
     cossinba = sina * cosb
     coscosba = cosb * cosa
     Form1.Cls
     NewPaint
End Sub
Sub rolldown()
     Iz_angle = (Iz_angle - 5)
     cosb = Cos(Iz_angle * pi180)
     sinb = Sin(Iz_angle * pi180)
     sinsinba = sinb * sina
     sincosba = sinb * cosa
     cossinba = sina * cosb
     coscosba = cosb * cosa
     Form1.Cls
     NewPaint
End Sub
Sub rollright()
     Ixy_angle = (Ixy_angle - 5)
     cosa = Cos(Ixy_angle * pi180)
     sina = Sin(Ixy_angle * pi180)
     sinsinba = sinb * sina
     sincosba = sinb * cosa
     cossinba = sina * cosb
     coscosba = cosb * cosa
     Form1.Cls
     NewPaint
End Sub
Sub rollleft()
     Ixy_angle = (Ixy_angle + 5)
     cosa = Cos(Ixy_angle * pi180)
     sina = Sin(Ixy_angle * pi180)
     sinsinba = sinb * sina
     sincosba = sinb * cosa
     cossinba = sina * cosb
     coscosba = cosb * cosa
     Form1.Cls
     NewPaint
End Sub
'This subroutine identifies the code of the pressed key
Private Sub Form_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
 Case 97
  ww = 1
 Case 100
  ww = 2
 Case 119
  ww = 3
 Case 120
  ww = 4
 Case 49
  ww = 5
 Case 50
  ww = 6
 Case 27
  Unload Me
 End Select
End Sub
Private Sub Form_Load()
 pi180 = 0.01745392
 Ixy_angle = 270
 Iz_angle = 85
 cosa = Cos(Ixy_angle * pi180)
 sina = Sin(Ixy_angle * pi180)
 cosb = Cos(Iz_angle * pi180)
 sinb = Sin(Iz_angle * pi180)
 sinsinba = sinb * sina
 sincosba = sinb * cosa
 cossinba = sina * cosb
 coscosba = cosb * cosa
 dYYshift = 80
 dXXshift = 80
 zoom = 6#
 NewPaint
End Sub
'This subroutine draws the cube using the translation code
Sub NewPaint()
 posxy -1, -1, -1: xxx = csx: yyy = csy:
 posxy -1, 1, -1: Line (xxx, yyy)-(csx, csy), QBColor(15): x = csx: y = csy
 posxy -1, 1, 1: Line (x, y)-(csx, csy), QBColor(15): x = csx: y = csy
 posxy -1, -1, 1: Line (x, y)-(csx, csy), QBColor(15): Line (csx, csy)-(xxx, yyy), QBColor(15)
 posxy 1, -1, -1: xxx = csx: yyy = csy:
 posxy 1, 1, -1: Line (xxx, yyy)-(csx, csy), QBColor(15): x = csx: y = csy
 posxy 1, 1, 1: Line (x, y)-(csx, csy), QBColor(15): x = csx: y = csy
 posxy 1, -1, 1: Line (x, y)-(csx, csy), QBColor(15): Line (csx, csy)-(xxx, yyy), QBColor(15)
 posxy 1, -1, -1: x = csx: y = csy: posxy -1, -1, -1: Line (x, y)-(csx, csy), QBColor(15)
 posxy 1, -1, 1: x = csx: y = csy: posxy -1, -1, 1: Line (x, y)-(csx, csy), QBColor(15)
 posxy 1, 1, 1: x = csx: y = csy: posxy -1, 1, 1: Line (x, y)-(csx, csy), QBColor(15)
 posxy 1, 1, -1: x = csx: y = csy: posxy -1, 1, -1: Line (x, y)-(csx, csy), QBColor(15)
End Sub
'This subroutine reads the value of the next rotation / zoom
Private Sub Timer1_Timer()
Select Case ww
Case 1
 rollleft
Case 2
 rollright
Case 3
 rollup
Case 4
 rolldown
Case 5
 zoom = zoom * 1.01
 Form1.Cls
 NewPaint
Case 6
 zoom = zoom * 0.99
 Form1.Cls
 NewPaint
End Select
End Sub
```

