VERSION 5.00
Begin VB.Form FormX 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flocking  Swarm AI - By Nathiel T. Tinsley"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FormX.frx":0000
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Start 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "FormX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Swarm AI - By Whisker
'©2002 Nazgul Software LLC

Option Explicit 'all variables must be dimensioned

Option Base 1 'set lower bound of arrays to 1

'Need to declare the timer so we can control the frame rate
Private Declare Function GetTickCount Lib "kernel32" () As Long

'fast way to set the color of a pixel
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Color As Long) As Byte

'used to keep the form on top of others
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOSIZE = &H1, SWP_NOMOVE = &H2, HWND_TOPMOST = -1

'creates the data structure for each particle
Private Type Particle
    xPos As Long 'current x position
    yPos As Long 'current y postition
    xVel As Long 'velocity on the x axis
    yVel As Long 'velocity on the y axis
    Index As Long 'index of collision partner
End Type

'creats data structure for average velocities
Private Type AverageVelocity
    Z As Long 'null counter
    P As Long 'positive counter
    N As Long 'negative counter
End Type

'creates the data array variables
Dim Bug() As Particle, Matrix() As Long

'creates the loop variables (don't dimension variables in sub routines if you can help it)
Dim Index As Long, Index1 As Long, Index2 As Long, I As Long, x As Long, y As Long

'creates misc storage variables (acsessing variables is much faster than object properties)
Dim FormH As Long, FormW As Long, FormHDC As Long, uBug As Long

'creates working variables
Dim CenterX As Long, CenterY As Long
Dim Neighbor As Boolean, Colision As Boolean 'state of scan area flags
Dim xAvgVel As AverageVelocity, yAvgVel As AverageVelocity 'average velocity counters
Dim LowX As Long, HighX As Long, LowY As Long, HighY As Long 'scan area bounds

Dim mintFollowX As Integer  'Where is the flock moving?
Dim mintFollowY As Integer

Dim mlngTimer As Long       'Timer for FPS maintenance
Dim mblnRunning As Boolean  'Render loop control variable

Const MS_PER_FRAME = 25     'How many milliseconds per frame of animation?
Const NUM_PARTICLES = 200   'How many particles will we use?
Const MAX_SPEED_VARIANCE = 0.01     'How fast can a sheep 'accelerate' w.r.t. his neighbour?
Const MIN_SPEED = 0.05      'Min particle velocity!
Const MAX_SPEED = 0.1       'Max particle velocity!
Const MIN_SEPERATION = 20   'Minimum distance between neighbouring particle
Const MAX_NOISE = 250       'Adds a little "jiggle" for realism
Const FOLLOW_AMOUNT = 100   'Speed with which flock will move when arrow keys are pressed

Const CIRCLE_RADIUS = 5     'Size of the circle that represents our sheep


'ensure unload
Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Start_Timer()

Dim I As Integer
Dim j As Integer
Dim blnSeperation As Boolean


'stayontop (3 times because win2k doesn't always see one or two)
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

'store form height, width, and hDC into fast variables
Me.ScaleHeight = 384
Me.ScaleWidth = 512
FormH = Me.ScaleHeight - 1
FormW = Me.ScaleWidth - 1
FormHDC = Me.hdc

'initialize particle array
ReDim Bug(NUM_PARTICLES - 1) 'set the number of bugs
uBug = UBound(Bug) 'place ubound of array into variable (speed up loops it is needed in)
Randomize Timer 'reseeds the random number generator in order to randomize locations.
For Index = 1 To uBug
    blnSeperation = False
    Do While Not (blnSeperation)
        '   Choose any location
            Bug(Index).xPos = Int(FormW * Rnd) + 1 'random x position
            Bug(Index).yPos = Int(FormH * Rnd) + 1 'random y position
            '   This is to ensure that it's not too close to a pre-existing particle.
            blnSeperation = True
            For j = 1 To Index - 1
                If CalcDist(Index, j) <= MIN_SEPERATION Then
                    blnSeperation = False
                    Exit For
                End If
            Next j
    Loop
Next Index

'initialize matrix array
ReDim Matrix(FormH, FormW) 'creates a matrix to store particle indexies
    
'Show the form
Me.Show

Call Main 'initialize main program loop

Start = False 'turn off start up timer

End Sub

'   Paint the postition of each particle on the form
Private Sub PaintIt()
On Error Resume Next
'   Draw our beautiful flock!
Me.Cls  '    Erase existing pixels
For Index = 1 To uBug
    SetPixelV FormHDC, Bug(Index).xPos, Bug(Index).yPos, vbGreen
    'Circle (Bug(Index).xPos, Bug(Index).yPos), CIRCLE_RADIUS, vbGreen
Next Index
End Sub

'   Main program loop (much faster than say, a timer)
Private Sub Main()
    
    'Start the main loop
    mblnRunning = True
    Do While mblnRunning
        'Maintain the FPS
        If mlngTimer + MS_PER_FRAME <= GetTickCount() Then
            'Reset the timer
            mlngTimer = GetTickCount()
            'Run the flocking AI simulation.
            Call Particles
            DoEvents
            'Display the particles.
            Call PaintIt
        End If
        'Let windows have a go
        DoEvents
    Loop
    
End Sub

Private Sub Particles()

Dim I As Long
Dim j As Long
Dim BugDist As Long
Dim BugXSpeed As Long
Dim BugYSpeed As Long
Dim BugXAvg As Long
Dim BugYAvg As Long
Dim BugXDist As Long
Dim BugYDist As Long
Dim BugTmpX As Long
Dim BugTmpY As Long

On Error Resume Next

'   Commence the particle loop, step through each particle...
For Index = 1 To uBug

    '   Find speed of nearest neighbour
    BugDist = 0
    For j = 1 To uBug
        '   Skip the current particle!
        If j <> Index Then
            '   Compare this particle's distance to the closest one so far.
            If (BugDist = 0) Or CalcDist(Index, j) < BugDist Then
                '   If this particle is closer, then store the distance.
                BugDist = CalcDist(Index, j)
                '   And store the speed
                BugXSpeed = Bug(j).xPos
                BugYSpeed = Bug(j).yPos
            End If
        End If
    Next j
        
    '   Let our particles move as fast as possible.
    Bug(Index).xVel = BugXSpeed + MAX_SPEED_VARIANCE
    Bug(Index).yVel = BugYSpeed + MAX_SPEED_VARIANCE
    If Bug(Index).xVel < MIN_SPEED Then Bug(Index).xVel = MIN_SPEED
    If Bug(Index).yVel < MIN_SPEED Then Bug(Index).yVel = MIN_SPEED
    If Bug(Index).xVel > MAX_SPEED Then Bug(Index).xVel = MAX_SPEED
    If Bug(Index).yVel > MAX_SPEED Then Bug(Index).yVel = MAX_SPEED

    '   Calculate and find the center of flock.
    CenterX = 0: CenterY = 0
    For I = 1 To uBug
        CenterX = CenterX + Bug(I).xPos
        CenterY = CenterY + Bug(I).yPos
    Next I
    CenterX = (CenterX / uBug) - 4
    CenterY = (CenterY / uBug)
        
    '   Average the values (and add some positive or negative noise and the "follow" amount).
    BugXAvg = (CenterX) + ((Rnd(MAX_NOISE) * MAX_NOISE) - (MAX_NOISE / 2)) + mintFollowX
    BugYAvg = (CenterY) + ((Rnd(MAX_NOISE) * MAX_NOISE) - (MAX_NOISE / 2)) + mintFollowY
           
    '   Set the scan area variables, cut off any 'past edge' area.
    LowX = Bug(Index).xPos - 4: If LowX < 1 Then LowX = 1
    LowY = Bug(Index).yPos - 4: If LowY < 1 Then LowY = 1
    HighX = Bug(Index).xPos + 4: If HighX > FormW Then HighX = FormW
    HighY = Bug(Index).yPos + 4: If HighY > FormH Then HighY = FormH
    
    '   Zero out the last velocity.
    Bug(Index).xVel = 0
    Bug(Index).yVel = 0
    
    '   Zero out the scan flags
    Neighbor = False
    Colision = False
    
    '   Scan the area for neighbors
    For x = LowX To HighX
    For y = LowY To HighY
        If Matrix(x, y) <> 0 And Matrix(x, y) <> Index Then
            
            '   set neighbor flag
            Neighbor = True
            
            '   if colision then set colision flag and colision index
            If Abs(Bug(Index).xPos - x) < 2 Or Abs(Bug(Index).yPos - y) < 2 Then
                Colision = True
                Bug(Index).Index = Matrix(x, y)
            End If
            
            '   add appropriate x axis velocity counter
            Select Case Bug(Matrix(x, y)).xVel
            Case 0
                xAvgVel.Z = xAvgVel.Z + 1
            Case Is > 0
                xAvgVel.P = xAvgVel.P + 1
            Case Is < 0
                xAvgVel.N = xAvgVel.N + 1
            End Select
            
            '   add appropriate y axis velocity counter
            Select Case Bug(Matrix(x, y)).yVel
            Case 0
                yAvgVel.Z = yAvgVel.Z + 1
            Case Is > 0
                yAvgVel.P = yAvgVel.P + 1
            Case Is < 0
                yAvgVel.N = yAvgVel.N + 1
            End Select
            
        End If
    Next y
    Next x
    
    '   if neighbor then move with average moan velocity
    If Neighbor = True And Colision = False Then
        '   calculate moan average velocity
        If xAvgVel.Z > xAvgVel.P And xAvgVel.Z > xAvgVel.N Then Bug(Index).xVel = 0
        If xAvgVel.P > xAvgVel.Z And xAvgVel.P > xAvgVel.N Then Bug(Index).xVel = 1
        If xAvgVel.N > xAvgVel.P And xAvgVel.N > xAvgVel.Z Then Bug(Index).xVel = -1
        If yAvgVel.Z > yAvgVel.P And yAvgVel.Z > yAvgVel.N Then Bug(Index).yVel = 0
        If yAvgVel.P > yAvgVel.Z And yAvgVel.P > yAvgVel.N Then Bug(Index).yVel = 1
        If yAvgVel.N > yAvgVel.P And yAvgVel.N > yAvgVel.Z Then Bug(Index).yVel = -1
        '   zero average velocity counters
        xAvgVel.Z = 0: xAvgVel.P = 0: xAvgVel.N = 0
        yAvgVel.Z = 0: yAvgVel.P = 0: yAvgVel.N = 0
    End If
    
    '   if collision then move away from colision
    If Colision = True Then
        Bug(Bug(Index).Index).xVel = Bug(Bug(Index).Index).xVel * -1
        Bug(Bug(Index).Index).yVel = Bug(Bug(Index).Index).yVel * -1
    End If
    
    '   if alone, then move towards center of flock
    If Neighbor = False Then
        If Bug(Index).xPos < BugXAvg Then Bug(Index).xVel = 1
        If Bug(Index).xPos > BugXAvg Then Bug(Index).xVel = -1
        If Bug(Index).xPos = BugXAvg Then Bug(Index).xVel = 0
        If Bug(Index).yPos < BugYAvg Then Bug(Index).yVel = 1
        If Bug(Index).yPos > BugYAvg Then Bug(Index).yVel = -1
        If Bug(Index).yPos = BugYAvg Then Bug(Index).yVel = 0
        
        '   Move towards the center! (as fast as allowable)
        BugTmpX = Bug(Index).xPos
        BugTmpY = Bug(Index).yPos
    
        '   Determine the X and Y movement amounts
        BugXDist = BugXAvg - Bug(Index).xPos
        BugYDist = BugYAvg - Bug(Index).yPos
        
         '  Move the X and Y coords
        Bug(Index).xPos = Bug(Bug(Index).xVel).xPos + BugXDist * Bug(Index).xVel
        Bug(Index).yPos = Bug(Bug(Index).yVel).yPos + BugYDist * Bug(Index).yVel
    
        '   Test for seperation
        For j = 0 To uBug
            If (Index <> j) And (CalcDist(Index, j) <= MIN_SEPERATION) Then
                '   There's another sheep too close, don't move
                Bug(Index).xPos = BugTmpX
                Bug(Index).yPos = BugTmpY
                Exit For
            End If
        Next j
    End If
    
    '   If no motion has been chosen, then move randomly.
    If Bug(Index).xVel = 0 And Bug(Index).yVel = 0 Then
        Bug(Index).xVel = Int(3 * Rnd) - 1
        Bug(Index).yVel = Int(3 * Rnd) - 1
    End If
    
    '   Wrap the paricles at the edges of the window; thus, correcting
    '   for out of boundaries.
    If Bug(Index).xPos > FormW Then
        Bug(Index).xVel = -1
        Bug(Index).xPos = Bug(Bug(Index).xVel).xPos - FormW
    End If
    If Bug(Index).xPos < 1 Then
        Bug(Index).xVel = 1
        Bug(Index).xPos = Bug(Bug(Index).xVel).xPos + FormW
    End If
    If Bug(Index).yPos > FormH Then
        Bug(Index).yVel = -1
        Bug(Index).yPos = Bug(Bug(Index).yVel).yPos - FormH
    End If
    If Bug(Index).yPos < 1 Then
        Bug(Index).yVel = 1
        Bug(Index).yPos = Bug(Bug(Index).yVel).yPos + FormH
    End If
        
    '   Remove old index from matrix
    Matrix(Bug(Index).xPos, Bug(Index).yPos) = 0
    
    '   Update particle position after new velocity
    Bug(Index).xPos = Bug(Index).xPos + Bug(Index).xVel
    Bug(Index).yPos = Bug(Index).yPos + Bug(Index).yVel
    
    '   Add new index to matrix.
    Matrix(Bug(Index).xPos, Bug(Index).yPos) = Index
    
Next Index

End Sub

Private Function CalcDist(ByVal intIndex1 As Integer, ByVal intIndex2 As Integer) As Long

    'How far appart are the two particles that have been indicated?
    CalcDist = Sqr((Bug(intIndex1).xPos - Bug(intIndex2).xPos) ^ 2 + (Bug(intIndex1).yPos - Bug(intIndex2).yPos) ^ 2)

End Function
