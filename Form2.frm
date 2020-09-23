VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controls"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3945
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pop-Up"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdScreen 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   3360
      Top             =   1200
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&START"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&PAUSE"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ListBox lstReport 
      Height          =   1815
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdApples 
      Caption         =   "&AUTO FEED"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "&KILL ANT!!"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&RESET"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblReport 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkPop_Click()

    If popup = True Then popup = False Else popup = True

End Sub

Private Sub cmdReset_Click()
          
    Call cmdClear_Click
    Call loadForm
       
End Sub

Private Sub cmdApples_Click()

    If Form1.Apples.Enabled = False Then Form1.Apples.Enabled = True Else Form1.Apples.Enabled = False

End Sub

Private Sub cmdClear_Click()

    Form1.Picture1.Cls
    
End Sub

Private Sub cmdKill_Click()
    Dim antKill As Integer
   
    DoEvents
    
    If lstReport.ListIndex = -1 Then
kill:
        DoEvents
        antKill = Int(Rnd() * antNum) + 1
        
        If Ant(antKill).IsDead = False Then
            Call kill_ant(antKill)
        
        Else:
            GoTo kill
        
        End If
        
        
    Else
        antKill = lstReport.ListIndex + 1
        
        If Ant(antKill).IsDead = False Then
            Call kill_ant(antKill)
        
        Else
            lblReport.Caption = UCase(Mid$(Ant(antKill).Name, 3, 3) & " IS ALREADY DEAD!!")
            Timer.Enabled = True
            Exit Sub
        
        End If
        
    End If
    
End Sub

Private Sub cmdPause_Click()
        
    If Timer.Enabled = True Then Timer.Enabled = False
    If Form1.Apples.Enabled = True Then Form1.Apples.Enabled = False
    
    lblReport.ForeColor = vbRed
    If pause = True Then
        Beep
        lblReport.Caption = ""
        pause = False
        Exit Sub
            
    Else
        Beep
        pause = True
        lblReport.Caption = "PAUSED"
        Exit Sub
        
    End If

End Sub

Private Sub cmdScreen_Click()
    
    With Form1
        .Width = Screen.Width
        .Height = Screen.Height + 400
        .Left = 0
        .Top = -400
        .BorderStyle = 0
    End With
            
    saver = True
    Form1.Apples.Enabled = True
    chkPop.Value = False
    Form2.Hide
    
    
End Sub

Private Sub cmdSearch_Click()
        
    cmdScreen.Enabled = True
    Form1.Picture1.Enabled = True
    cmdSearch.Enabled = False
    Run = True
    'The main loop for making the ants run about
    Do
        DoEvents
    
        If Run = False Then cmdSearch.Enabled = True: Exit Do
        For loopval = 1 To antNum
        
                Call gotoPoint(Ant(loopval))
                                
        Next loopval
    
    Loop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    popup = True

End Sub

Private Sub Timer_Timer()

    lblReport.Caption = ""
    Timer.Enabled = False

End Sub


