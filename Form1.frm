VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANTSIM!"
   ClientHeight    =   5430
   ClientLeft      =   1155
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgApple 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   450
      Left            =   240
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox imgBlack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   450
      Left            =   240
      Picture         =   "Form1.frx":09DC
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox imgVert 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   450
      Left            =   600
      Picture         =   "Form1.frx":0A9E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox imgHorz 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      Picture         =   "Form1.frx":0B60
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox imgAntDeadVert 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   450
      Left            =   600
      Picture         =   "Form1.frx":0BF2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox imgAntDeadHorz 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      Picture         =   "Form1.frx":12C4
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox imgAntRight 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   120
      Left            =   240
      Picture         =   "Form1.frx":197E
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox imgAntLeft 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   120
      Left            =   240
      Picture         =   "Form1.frx":19E8
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox imgAntUp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      Picture         =   "Form1.frx":1A52
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox imgAntDown 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      Picture         =   "Form1.frx":1AD0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   5415
      Left            =   0
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.Timer Apples 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   240
         Top             =   2760
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'used to slow it down when there are few ants

Private Sub Form_Load()
    
    Call Randomize
  
    'Array storing the possible names of the ants
    names(1) = "Bob"
    names(2) = "Jim"
    names(3) = "Dan"
    names(4) = "Tim"
    names(5) = "Baz"
    names(6) = "Reg"
    names(7) = "Tom"
    names(8) = "Rob"
    names(9) = "Joe"
    names(10) = "Marmaduke"
        
    Limit = 0
    
    Me.Show
    DoEvents
 
    Call loadForm

End Sub

Private Sub Form_Resize()

    Picture1.Width = ScaleX(Form1.Width, vbTwips, vbPixels) - 5
    Picture1.Height = ScaleY(Form1.Height, vbTwips, vbPixels) - 5
    Picture1.Left = 0
    Picture1.Top = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
            
    Unload Form2
    End

End Sub





Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim AppleX As Integer
    Dim AppleY As Integer
    Dim IsAnt As Boolean
    Dim antCount As Integer
            
    If saver = True And Button = 1 Then
        saver = False
        Form2.Show
        Form2.ZOrder 1
        
        With Form1
            .Width = 6650
            .Height = 6650
            .Top = Screen.Height / 4
            .Left = Screen.Width / 4
        End With
    
    End If
        
    If popup = True Then Form2.Show
    antCount = 0
    IsAnt = False
        
    'BitBlt Picture1.hDC, X - 90, Y - 90, Picture2.Width, Picture2.Height, Picture2.hDC, 0, 0, vbMergePaint
    'BitBlt Picture1.hDC, X - 90, Y - 90, Picture2.Width, Picture2.Height, Picture2.hDC, 0, 0, vbSrcAnd
    'Sleep (100)
    'BitBlt Picture1.hDC, X - 90, Y - 90, Picture2.Width, Picture2.Height, Picture2.hDC, 0, 0, vbMergePaint

        
        'Adjust the click co-ords to make the apple appear under the
        'cursors hotspot
        If 0 < X < Picture1.Width Then AppleX = X - 10 Else GoTo endstop
        If 0 < Y < Picture1.Height Then AppleY = Y - 15 Else GoTo endstop
            
        
        'Clear the way then draw an apple
        BitBlt Picture1.hDC, AppleX, AppleY, imgBlack.Width, imgBlack.Height, imgBlack.hDC, 0, 0, vbMergePaint
        BitBlt Picture1.hDC, AppleX, AppleY, imgApple.Width, imgBlack.Height, imgApple.hDC, 0, 0, vbSrcAnd
        
        'This part randomly sets ants to eat the apple if they aren't already
        
        For loopval = 1 To antNum
            
            If Ant(loopval).DestType <> DestApple And Ant(loopval).IsDead = False Then
                If Sqr((Abs(AppleY - Ant(loopval).Y)) ^ 2 + (Abs(AppleX - Ant(loopval).X)) ^ 2) < 90 Then
                    With Ant(loopval)
                        .Destination.X = AppleX
                        .Destination.Y = AppleY
                        .DestType = DestApple
                    End With
                
                    IsAnt = True
                    antCount = antCount + 1
                    If antCount = antNum / 5 Then Exit For
                
                End If
            End If
            
        Next loopval



endstop:
    
    'If no ants have accepted the apple then blank it out and display
    'a message
    
    If IsAnt = False Then
        DoEvents
        BitBlt Picture1.hDC, AppleX, AppleY, imgBlack.Width, imgBlack.Height, imgBlack.hDC, 0, 0, vbMergePaint
        Form2.lblReport.ForeColor = vbRed
        Form2.lblReport.Caption = "Bad Apple! REJECTED!!!"
        Form2.Timer.Enabled = True
    End If


End Sub

'@@@@@@@@@@@@@@
'@PRIVATE SUBS@
'@@@@@@@@@@@@@@



'@@@@@@@
'@TIMER@
'@@@@@@@

Private Sub Apples_Timer()

    Call Picture1_MouseDown(2, 0, Int(Rnd() * (Picture1.Width - 20)) + 10, Int(Rnd() * (Picture1.Height - 20)) + 10)

End Sub

