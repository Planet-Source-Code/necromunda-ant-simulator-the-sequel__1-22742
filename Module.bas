Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Global Const antNum = 250 'Changing this changes the number of ants/apples
Global Const sleepTime = 0

'This is used for the randomly selected point
'that each ant goes to. i originally used random movement
'but it just made them dance about

Global saver As Boolean
Global popup As Boolean

Type pointXY
    X As Long
    Y As Long
End Type

Type Ant
    Name As String
    X As Integer    'Coordinates
    Y As Integer
    Direct As Integer 'Which way its facing
    Destination As pointXY 'Where its going
    DestType As Integer 'What it's going to
    ApplesEaten As Integer 'How much its eaten
    IsDead As Boolean 'Is it dead?
End Type

Global Ant(antNum) As Ant

Global Const DestApple = 0     'Constants for destination type
Global Const DestGotoPoint = 1


Global pause As Boolean
Global random As Integer

Global Limit As Integer
Global KilledCount As Integer

Global loopval As Long

Global Const goUp = 1      'constants for each direction
Global Const goDown = 2
Global Const goLeft = 3
Global Const goRight = 4

Global names(10) As String
Global Run As Boolean

Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long) 'Sleeeeeep
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Sub loadForm()
    'Setup the ants
        
    For loopval = 1 To antNum
        'Firstly, setup a Goto point
        Ant(loopval).X = Int(Rnd() * Form1.Picture1.Width)
        Ant(loopval).Y = Int(Rnd() * Form1.Picture1.Height)
                
        'Assign a name and reset other variables
        Ant(loopval).Name = names(Int(Rnd() * 10) + 1)
        Ant(loopval).ApplesEaten = 0
        Ant(loopval).IsDead = False
                    
        'Create the random point for the ant to go to
        Call createPoint(Ant(loopval))
            
        'Then Draw the ant, and setup the graphical variables
        BitBlt Form1.Picture1.hDC, Ant(loopval).X, Ant(loopval).Y, Form1.imgAntUp.Width, Form1.imgAntUp.Height, Form1.imgAntUp.hDC, 0, 0, vbSrcAnd
        Ant(loopval).Direct = 1
        Ant(loopval).DestType = DestGotoPoint
    Next loopval
    
    'Fill the scoreboard listbox
    Form2.lstReport.Clear
    For loopval = 1 To antNum
        Form2.lstReport.AddItem Ant(loopval).Name & Space(2) & Ant(loopval).ApplesEaten
    
    Next loopval
    
    Form2.Show

End Sub


 Sub kill_ant(number As Integer)
    
    'First set the isdead bool variable to true
    Ant(number).IsDead = True
            
    'Determine the direction it's facing (horizontal or vertical) and
    'draw the dead image.
                
    If Ant(number).Direct = 3 Or Ant(number).Direct = 4 Then
        BitBlt Form1.Picture1.hDC, Ant(number).X - 5, Ant(number).Y - 5, Form1.imgHorz.Width, Form1.imgHorz.Height, Form1.imgHorz.hDC, 0, 0, vbMergePaint
        BitBlt Form1.Picture1.hDC, Ant(number).X - 5, Ant(number).Y - 5, Form1.imgAntDeadHorz.Width, Form1.imgAntDeadHorz.Height, Form1.imgAntDeadHorz.hDC, 0, 0, vbSrcAnd
       
    ElseIf Ant(number).Direct = 1 Or Ant(number).Direct = 2 Then
        BitBlt Form1.Picture1.hDC, Ant(number).X - 5, Ant(number).Y - 5, Form1.imgVert.Width, Form1.imgVert.Height, Form1.imgVert.hDC, 0, 0, vbMergePaint
        BitBlt Form1.Picture1.hDC, Ant(number).X - 5, Ant(number).Y - 5, Form1.imgAntDeadVert.Width, Form1.imgAntDeadVert.Height, Form1.imgAntDeadVert.hDC, 0, 0, vbSrcAnd
    End If
        
    'Display dead message
    Form2.lblReport.ForeColor = vbRed
    Form2.lblReport.Caption = Ant(number).Name & " KILLED!!"
    Ant(number).Name = "xx" & Ant(number).Name & "xx"
    
    Form2.Timer.Enabled = True
    
    Form2.lstReport.Clear
    For loopval = 1 To antNum
        Form2.lstReport.AddItem Ant(loopval).Name & Space(2) & Ant(loopval).ApplesEaten
    
    Next loopval

    
End Sub
 
 
 Sub gotoPoint(Ant As Ant)
'This is for searching out the random point
    
    'Pause:
    'If pause has been pressed then loop until it's not
    
    If pause = True Then
        Do
            DoEvents
        Loop Until pause = False

    End If
    
    If Ant.IsDead = True Then
       If Ant.Direct = 3 Or Ant.Direct = 4 Then
            BitBlt Form1.Picture1.hDC, Ant.X - 5, Ant.Y - 5, Form1.imgAntDeadHorz.Width, Form1.imgAntDeadHorz.Height, Form1.imgAntDeadHorz.hDC, 0, 0, vbSrcAnd
            Exit Sub
        
        ElseIf Ant.Direct = 1 Or Ant.Direct = 2 Then
            BitBlt Form1.Picture1.hDC, Ant.X - 5, Ant.Y - 5, Form1.imgAntDeadVert.Width, Form1.imgAntDeadVert.Height, Form1.imgAntDeadVert.hDC, 0, 0, vbSrcAnd
            Exit Sub
            
        End If
    
    End If
        
    'If its got there and it's a goto point then create a new one
    If Ant.X = Ant.Destination.X And Ant.Y = Ant.Destination.Y And Ant.DestType = DestGotoPoint Then
        Call createPoint(Ant)
        Exit Sub
    
    'If it gets there and it's an apple point then eat it.
    ElseIf Ant.X = Ant.Destination.X And Ant.Y = Ant.Destination.Y And Ant.DestType = DestApple Then
        Call appleEat(Ant)
        Exit Sub
    
    End If

    'Makes the movement seem more natural
    'it randomly chooses between going up/down or left/right
    
    If Int(Rnd() * 2) = 1 Then
        If Ant.Y < Ant.Destination.Y Then
            Call moveAnt(Ant, goUp)
        ElseIf Ant.Y > Ant.Destination.Y Then
            Call moveAnt(Ant, goDown)
        End If
        Exit Sub
    
    Else
        If Ant.X < Ant.Destination.X Then
            Call moveAnt(Ant, goRight)
        ElseIf Ant.X > Ant.Destination.X Then
            Call moveAnt(Ant, goLeft)
        End If
        Exit Sub
    
    End If
        
End Sub

 Sub createPoint(Ant As Ant)
       
    'Create the X and Y co-ords for the ant to travel to
    
    Ant.Destination.X = 5 + Int(Rnd() * Form1.Picture1.Width - 5)
    Ant.Destination.Y = 5 + Int(Rnd() * Form1.Picture1.Height - 5)
    
    'Set the destination type as a goto point
    Ant.DestType = DestGotoPoint
    
End Sub


 Sub moveAnt(Ant As Ant, direction As Integer)
    
    Sleep (sleepTime)
    'This slows down the process
    
    'Don't move the ant if it's dead.
    If Ant.IsDead = True Then Exit Sub
    
    'This section demonstrates the use of BitBlt...(kinda)
    
    'First it goes through this bit below, which draws an ant
    'but reversed.  It draws only the black bits, but as white.
    'This clears the ant image out of the way before drawing the new
    'one.
    
    
    'BitBlt: A Rough Guide
    'hDestDC = the Hdc (eg. form1.picture1.hdc) of the picture box where
    'you want to copy the image
    
    'X = Where to draw in X
    'Y = Where to draw in Y
    
    'nWidth = The width of the image being draw
    'nHeight = The height
    
    'hSrcDC = The hdc of the source picture box
    
    'xSrc = where on the source pic box X
    'ySrc = where on the source pic box Y
    
    'dwRop = how to copy the image
    
        'vbMergePaint = copy only black bits (as white)
        'vbSrcAnd = copy non white bits
        'vbNotSrcAnd = Inverse
        'vbSrcCopy = Copies image
    
    '(This wont work with image boxes as they dont have an hDC)
        
    
    Select Case Ant.Direct
        Case Is = 1     'Go UP
            BitBlt Form1.Picture1.hDC, Ant.X, Ant.Y, Form1.imgAntUp.Width, Form1.imgAntUp.Height, Form1.imgAntUp.hDC, 0, 0, vbMergePaint
        
        Case Is = 2     'Go DOWN
            BitBlt Form1.Picture1.hDC, Ant.X, Ant.Y, Form1.imgAntDown.Width, Form1.imgAntDown.Height, Form1.imgAntDown.hDC, 0, 0, vbMergePaint
        
        Case Is = 3     'Go LEFT
            BitBlt Form1.Picture1.hDC, Ant.X, Ant.Y, Form1.imgAntLeft.Width, Form1.imgAntLeft.Height, Form1.imgAntLeft.hDC, 0, 0, vbMergePaint
        
        Case Is = 4     'Go RIGHT
            BitBlt Form1.Picture1.hDC, Ant.X, Ant.Y, Form1.imgAntRight.Width, Form1.imgAntRight.Height, Form1.imgAntRight.hDC, 0, 0, vbMergePaint
        
    End Select
    
    
    Select Case direction
        Case Is = goUp
            Ant.Y = Ant.Y + 1
            'If Ant.Y < Form1.Picture1.Height Then Ant.Y = Ant.Y - 1 Else Ant.Y = Ant.Y + 1
            BitBlt Form1.Picture1.hDC, Ant.X, Ant.Y, Form1.imgAntUp.Width, Form1.imgAntUp.Height, Form1.imgAntUp.hDC, 0, 0, vbSrcAnd
            Ant.Direct = 1
        Case Is = goDown
            Ant.Y = Ant.Y - 1
            'If Ant.Y > 0 Then Ant.Y = Ant.Y + 1 Else Ant.Y = Ant.Y - 1
            BitBlt Form1.Picture1.hDC, Ant.X, Ant.Y, Form1.imgAntDown.Width, Form1.imgAntDown.Height, Form1.imgAntDown.hDC, 0, 0, vbSrcAnd
            Ant.Direct = 2
        Case Is = goLeft
            Ant.X = Ant.X - 1
            'If Ant.X > 0 Then Ant.X = Ant.X + 1 Else Ant.X = Ant.X - 1
            BitBlt Form1.Picture1.hDC, Ant.X, Ant.Y, Form1.imgAntLeft.Width, Form1.imgAntLeft.Height, Form1.imgAntLeft.hDC, 0, 0, vbSrcAnd
            Ant.Direct = 3
        Case Is = goRight
            Ant.X = Ant.X + 1
            'If Ant.X < Form1.Picture1.Width Then Ant.X = Ant.X - 1 Else Ant.X = Ant.X + 1
            BitBlt Form1.Picture1.hDC, Ant.X, Ant.Y, Form1.imgAntRight.Width, Form1.imgAntRight.Height, Form1.imgAntRight.hDC, 0, 0, vbSrcAnd
            Ant.Direct = 4
    End Select
    
    DoEvents
    
End Sub


 Sub appleEat(AnAnt As Ant)
    'The movement to "eat" the apple
    Dim loop2 As Integer
            
    Form1.Picture1.Enabled = False
        
    'Add to the apple tally
    AnAnt.ApplesEaten = AnAnt.ApplesEaten + 1
    
    Form2.lblReport.ForeColor = vbGreen
    Form2.lblReport.Caption = "Eaten by " & AnAnt.Name
    'lblReport.ForeColor = vbRed
    Form2.Timer.Enabled = True
    
    For loopval = 1 To antNum
        If Ant(loopval).DestType = DestApple And _
            Ant(loopval).Destination.X = AnAnt.Destination.X And _
            Ant(loopval).Destination.Y = AnAnt.Destination.Y Then
            Ant(loopval).DestType = DestGotoPoint
            
        End If
    Next loopval
        
    For loopval = 1 To 5
        
        For loop2 = 1 To 12
            Call moveAnt(AnAnt, goRight)
        
        Next loop2
        Sleep (10)
        
        For loop2 = 1 To 5
            Call moveAnt(AnAnt, goUp)
        
        Next loop2
        Sleep (10)
        
        For loop2 = 1 To 13
            Call moveAnt(AnAnt, goLeft)
        
        Next loop2
        Sleep (10)
    Next loopval
    
    Form1.Picture1.Enabled = True
    
    Form2.lstReport.Clear
    
    For loopval = 1 To antNum
        Form2.lstReport.AddItem Ant(loopval).Name & Space(2) & Ant(loopval).ApplesEaten
    
    Next loopval
    
    Call createPoint(AnAnt)
    
    
End Sub

