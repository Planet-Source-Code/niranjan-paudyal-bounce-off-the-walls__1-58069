VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bounce off the walls!"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pw 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   6195
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   409
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   394
      TabIndex        =   0
      Top             =   0
      Width           =   5970
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Bounce off the walls!
'Created by Niranjan Paudyal (nirpaudyal@hotmail.com) 2/1/2005
'Please feel free to use this code as you wish!
'Dont understand something, than contact me!

'I am sure you can find a 100 ways to make this code more faster as I have tried to
'keep everything simple and 'easy' to understand
'Enjoy!

'HOW TO USE THE PROGRAM--------------------------------
'WHEN YOU FIRST RUN IT, CLICK WITH YOU LEFT MOUSE BUTON
'WHERE YOU WANT THE BALL TO START MOVING FROM
'IT WILL THEN START BOUNCING ABOUT LIKE A IDIOT
'CLICK ANOTHER POINT AT ANYTIME TO GET THE BALL MOVING FROM THERE

'THE PICTURE ON THE PICTURE BOX CAN BE ANY COLOUR OR SHAPE, BUT THE AREA
'WHERE THE BALL CAN MOVE HAS TO BE BLACK!

'Limitations of the code
'The WALLS on the picture box has to be greater than or equal to a thickness of 2 pixels!
'speed of ball has to be less than or equal to the diameter of the ball!
'The bigger the ball, the more accurate the motion!
'----------------------------------------------------------------



Private Const Pi = 3.14159265358979
Private Type PointAPI
    X As Long
    Y As Long
End Type
Private Type PointSNG
    X As Single
    Y As Single
End Type
Private Type Ball
    Position As PointSNG
    Vel As PointSNG
    Radius As Long
    Mass As Long
End Type

Dim B As Ball

Dim Running As Boolean 'Is the program running?

Private Sub ChangeVelocities(B As Ball, cX As Long, cY As Long)
'ChangeVelocities(the ball in question, X position of contact with the wall, Y position of contact with the wall)
'This sub deals with the bouncing of ball off the wall
'This sub has been taken and modified from my 'bouncing balls' program
'The momemtum conservation is basically the same, however, in this program, walls
'have infinate mass and a velocity of 0, therefor, the wall will not move!
'See my 'bouncing balls' program as this procedure will make more sense then
    Dim X1 As Single, Y1 As Single
    Dim X2 As Single, Y2 As Single
    Dim angle As Single

    X1 = B.Position.X   'center X of the ball
    Y1 = B.Position.Y   'center Y of the ball
    X2 = cX 'X point of collision with wall
    Y2 = cY 'Y point of collision with wall
    
    'Get the angel between the ball and the wall
    If (X2 - X1) <> 0 Then angle = Atn((Y2 - Y1) / (X2 - X1)) Else angle = Pi / 2
    
    hX1 = B.Vel.X
    hY1 = B.Vel.Y
    hX2 = 0 'This is the velocity of the wall at point of contact, note it is 0! thats because the walls are not moving!
    hY2 = 0
    
    'resolve the velocitis such that they are along the line of collision
    X1 = hX1 * Cos(-angle) - hY1 * Sin(-angle)
    Y1 = hX1 * Sin(-angle) + hY1 * Cos(-angle)
    X2 = hX2 * Cos(-angle) - hY2 * Sin(-angle)
    'Y2 = hX2 * Sin(-angle) + hY2 * Cos(-angle)     'Left over from the Ball collision program, not needed here
    
    Mass = 1000000000    'This is the mass of the wall, otherewise, the balls energy will be lost to the wall and the ball will lose it velocity 'Try it by setting it to 100!
    'Momemtum is conserved in the line of collision
    hX1 = (X1 * (B.Mass - Mass) + (X2 * 2 * Mass)) / (B.Mass + Mass)
    hX2 = ((X1 * 2 * Mass) + X2 * (B.Mass - Mass)) / (B.Mass + Mass)
    
    'keep the vertical component in the line of collision remains the same
    hY1 = Y1
    'hY2 = Y2   'This is for the wall, so ignore it!
    
    'resolve back the velocities to their normal coordinates
    X1 = hX1 * Cos(angle) - hY1 * Sin(angle)    'For the ball
    Y1 = hX1 * Sin(angle) + hY1 * Cos(angle)
    'X2 = hX2 * Cos(angle) - hY2 * Sin(angle)   'For the wall
    'Y2 = hX2 * Sin(angle) + hY2 * Cos(angle)
    
    'set the velocitie of the ball
    B.Vel.X = X1
    B.Vel.Y = Y1
End Sub


Private Function IsTouchingWall(BX As Long, BY As Long, BRadius As Long, ByRef rx As Long, ByRef ry As Long) As Boolean
'IsToughingWall(X position of the ball,Y position of the ball,Radius of the ball, return the X point of contact with wall, return the Y point of contact with wall)
'This sub is used to identify if the ball has crashed with the wall
'It will fo through every point at the radius of the ball and compare the color of
'those points to the color on the picture box, Pw
'if the color is not 0 (black), the ball must have crashed with the wall at that point
'For this reason, the walls must be of color other than 0
'Note that the sub will automatically detect outside the picture box because the color there is set to -1 by windows
    
'If you want to speed this sub up then i suggest you use dibs to find the pixel color at the point
'I use Point because it makes the whole thing 'easier' to understand
    Dim X As Long, Y As Long
    Dim C As Long, br As Long
    
    br = BRadius * BRadius  'The hypotnuse^2 of the circle of the ball
    
    For X = 0 To BRadius
        Y = Sqr(br - X * X)
        C = Pw.Point(BX + X, BY + Y)
        If C <> 0 Then
            rx = BX + X: ry = BY + Y
            IsTouchingWall = True
            Exit Function
        End If
        C = Pw.Point(BX + X, BY - Y)
        If C <> 0 Then
            rx = BX + X: ry = BY - Y
            IsTouchingWall = True
            Exit Function
        End If
        C = Pw.Point(BX - X, BY + Y)
        If C <> 0 Then
            rx = BX - X: ry = BY + Y
            IsTouchingWall = True
            Exit Function
        End If
        C = Pw.Point(BX - X, BY - Y)
        If C <> 0 Then
            rx = BX - X: ry = BY - Y
            IsTouchingWall = True
            Exit Function
        End If
    Next X
    
    For Y = 0 To BRadius
        X = Sqr(br - Y * Y)
        C = Pw.Point(BX + X, BY + Y)
        If C <> 0 Then
            rx = BX + X: ry = BY + Y
            IsTouchingWall = True
            Exit Function
        End If
        C = Pw.Point(BX + X, BY - Y)
        If C <> 0 Then
            rx = BX + X: ry = BY - Y
            IsTouchingWall = True
            Exit Function
        End If
        C = Pw.Point(BX - X, BY + Y)
        If C <> 0 Then
            rx = BX - X: ry = BY + Y
            IsTouchingWall = True
            Exit Function
        End If
        C = Pw.Point(BX - X, BY - Y)
        If C <> 0 Then
            rx = BX - X: ry = BY - Y
            IsTouchingWall = True
            Exit Function
        End If
    Next Y
End Function
Private Sub GoBack(R As PointAPI)
'GoBack(Returen R is the point at which the actual collision occured)
    'This sub is called when there is a collision
    'If there is no collision and this sub is called, then there there will be problems!
    'The purpose of this sub is the seperate the ball from the wall.
    'if the ball is travelling at high speed, the ball will go into the wall, this
    'Will cause problems when doing the momemtum calculations as the ball will tend to slide along the wall rather than bounce
    'In order to solve this problem, 2 things have to be done
    '1) ball needs to be backtracked to find when its actual point of contact with the wall was
    '2) ball needs to be seperated from the wall
    'To achieve this, we go from the current point of the ball, backwards along the path it came
    'and location at which it no longer collides. The point just before this location is the point of contact!
    Dim LastIntersect As PointAPI
    Dim CurrentPoint As PointAPI
    
    'This little If, else is used to find how much a change in  X location of the ball affects the Y location by
    If Abs(B.Vel.Y) >= Abs(B.Vel.X) Then
        vs = 1
        hs = Abs(B.Vel.X) / Abs(B.Vel.Y)
    Else
        hs = 1
        vs = Abs(B.Vel.Y) / Abs(B.Vel.X)
    End If
    If B.Vel.Y > 0 Then vs = -vs
    If B.Vel.X > 0 Then hs = -hs
    
    Do
        'Update the position to check for collision by the above factors each time
        CurrentPoint.X = B.Position.X + hs * i
        CurrentPoint.Y = B.Position.Y + vs * i
        If IsTouchingWall(CurrentPoint.X, CurrentPoint.Y, B.Radius, LastIntersect.X, LastIntersect.Y) Then
            'If there is still collision than update the R value
            R.X = LastIntersect.X
            R.Y = LastIntersect.Y
        Else
            'If there is no more collisions, the sub must be exited, Remember that the R value will be returend from this procedure
            B.Position.X = B.Position.X + hs * (i + 1)  'This is used to reset the position of the ball to a point just after when it would have first made contact
            B.Position.Y = B.Position.Y + vs * (i + 1)  'Just after because sometimes it slides along the wall when it set to a point at which it just made contact!
            Exit Do
        End If
                
        'If after a certain ammount of going back, we cant find the point at which it first made contact, there must have been some error, so just quit the sub
        'otherwise, we get stuck in a loop!
        If i > B.Radius * 2 Then
            Exit Sub
        End If
        
        'Update the factor
        i = i + 1
    Loop
    
End Sub
Private Sub run()
    'The heart of the program
    
    Dim Xr As Long, Yr As Long
    Dim RR As PointAPI
    'Check to see if program is running
    While Running
        Pw.Cls
        'is there a collision?
        If IsTouchingWall(CSng(B.Position.X), CSng(B.Position.Y), B.Radius, Xr, Yr) Then
            'Backtrack to find actual point of collision
            GoBack RR
            'Change the speed of the ball
            ChangeVelocities B, RR.X, RR.Y
            
        End If
        'Draw the ball
        Pw.Circle (B.Position.X, B.Position.Y), B.Radius, vbGreen
        'Update the position of the ball
        B.Position.X = B.Position.X + B.Vel.X
        B.Position.Y = B.Position.Y + B.Vel.Y
        
        'allow time for other windows events
        DoEvents
    Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Running = False
End Sub

Private Sub Pw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This sub is used to set the starting properties of the ball and get is moving!
    Dim rx As Long, ry As Long
    B.Radius = 10
    If IsTouchingWall(CLng(X), CLng(Y), B.Radius, rx, ry) Then
        Exit Sub 'If the point at which the ball has been palced produces a collision, then dont bother settin the ball there!
    Else
        B.Mass = 2
        B.Position.X = X: B.Position.Y = Y
        B.Vel.Y = 3: B.Vel.X = 0
        Running = True
        run
    End If
End Sub
