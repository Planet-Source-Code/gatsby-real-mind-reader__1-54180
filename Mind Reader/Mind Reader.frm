VERSION 5.00
Begin VB.Form mindReader 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Benjamin Blasen's Mind Reader"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12735
   Icon            =   "Mind Reader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrClosed 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   6600
   End
   Begin VB.Timer tmrBlink 
      Interval        =   4000
      Left            =   480
      Top             =   6480
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   600
      MousePointer    =   2  'Cross
      OLEDropMode     =   2  'Automatic
      Picture         =   "Mind Reader.frx":0442
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   3105
   End
   Begin VB.Image buttonDown 
      Height          =   825
      Left            =   600
      MousePointer    =   2  'Cross
      OLEDropMode     =   2  'Automatic
      Picture         =   "Mind Reader.frx":0A86
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Mind Reader.frx":11C7
      Stretch         =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   7530
   End
   Begin VB.Image eye 
      Height          =   2535
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Mind Reader.frx":B913
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   7530
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   9
      Left            =   11400
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   8
      Left            =   10080
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   7
      Left            =   8880
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   6
      Left            =   7680
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   5
      Left            =   6480
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   4
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   3
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label display 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here When Ready"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Index           =   1
      X1              =   4200
      X2              =   4200
      Y1              =   4560
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Index           =   0
      X1              =   0
      X2              =   12720
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "mindReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Benjamin Blasen
' email any comments/suggestions to bblasen@hotmail.com

Option Explicit
Dim myAnswer As String
Dim letter(1 To 99) As String
Dim X As Integer
Dim newLetter As Integer
Dim arrayLetter As String
Dim message(1 To 10) As String
Dim blink As Boolean

Private Sub eye_Click()
    
    MsgBox "Pick any two digit number." + vbCrLf + "Add the two digits together and subtract the sum from the original number." + vbCrLf + "Once you have the answer, find the letter correlating to your number and stare at it for a few seconds." + vbCrLf + "When ready, click on the red button.", vbInformation, "Mind Reader Instructions"
    
End Sub

Private Sub Form_GotFocus()
    
    If myRefresh = True Then
    
        doRun
        myRefresh = False
        
    End If
    
End Sub

Private Sub Form_Load()

    doRun
    
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
   Image1.Visible = False
   buttonDown.Visible = True
   
End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
   Image1.Visible = True
   buttonDown.Visible = False
   myLetter = myAnswer
   answer.Show
    
End Sub

Private Function doRun()
    
    X = 1
    Do While X < 11 ' clear message array
   
        message(X) = ""
        X = X + 1
        
    Loop
    
    X = 1
    
    Do While X < 100   ' build new array of letters
        
        Randomize
        newLetter = Rnd * (15) ' Random multiplied be the number of possiblilities
        
        returnLetter (newLetter)
        
        letter(X) = arrayLetter
        
        X = X + 1
    
    Loop
    
    ' find and insert myAnser
    
    myAnswer = letter(9)
    myLetter = myAnswer
    letter(18) = myAnswer
    letter(27) = myAnswer
    letter(36) = myAnswer
    letter(45) = myAnswer
    letter(54) = myAnswer
    letter(63) = myAnswer
    letter(72) = myAnswer
    letter(81) = myAnswer
    
    X = 1
    
    Do While X < 100  ' build new message arrays
        
        If X < 11 Then
        
            message(1) = message(1) & X & ": " & letter(X) & vbCrLf
        
        ElseIf X > 10 And X < 21 Then
        
            message(2) = message(2) & X & ": " & letter(X) & vbCrLf
         
        ElseIf X > 20 And X < 31 Then
        
            message(3) = message(3) & X & ": " & letter(X) & vbCrLf
        
        ElseIf X > 30 And X < 41 Then
        
            message(4) = message(4) & X & ": " & letter(X) & vbCrLf
        
        ElseIf X > 40 And X < 51 Then
        
            message(5) = message(5) & X & ": " & letter(X) & vbCrLf
        
        ElseIf X > 50 And X < 61 Then
        
            message(6) = message(6) & X & ": " & letter(X) & vbCrLf
            
        ElseIf X > 60 And X < 71 Then
        
            message(7) = message(7) & X & ": " & letter(X) & vbCrLf
            
        ElseIf X > 70 And X < 81 Then
        
            message(8) = message(8) & X & ": " & letter(X) & vbCrLf
            
        ElseIf X > 80 And X < 91 Then
        
            message(9) = message(9) & X & ": " & letter(X) & vbCrLf
        
        ElseIf X > 90 And X < 100 Then
        
            message(10) = message(10) & X & ": " & letter(X) & vbCrLf
        
        End If
        
        X = X + 1
            
    Loop
    
    Dim myCount As Integer
    myCount = 0
    
    Do While myCount < 10
    
        display(myCount) = message(myCount + 1)
        myCount = myCount + 1
        
    Loop
    
End Function

Private Function returnLetter(myNum As Integer)

    If myNum = 1 Then
        
        arrayLetter = "A"
        
    ElseIf myNum = 2 Then
    
        arrayLetter = "E"
       
    ElseIf myNum = 3 Then
    
        arrayLetter = "I"
        
    ElseIf myNum = 4 Then
    
        arrayLetter = "O"
    
    ElseIf myNum = 5 Then
    
        arrayLetter = "U"
        
    ElseIf myNum = 6 Then
    
        arrayLetter = "X"
        
    ElseIf myNum = 7 Then
    
        arrayLetter = "Y"
        
    ElseIf myNum = 8 Then
    
        arrayLetter = "Z"
        
    ElseIf myNum = 9 Then
    
        arrayLetter = "L"
        
    ElseIf myNum = 10 Then
    
        arrayLetter = "R"
        
    ElseIf myNum = 11 Then
    
        arrayLetter = "Q"
        
    ElseIf myNum = 12 Then
    
        arrayLetter = "J"
        
    ElseIf myNum = 13 Then ' new
    
        arrayLetter = "F"
        
    ElseIf myNum = 14 Then
    
        arrayLetter = "B"
        
    ElseIf myNum = 15 Then
    
        arrayLetter = "W"
        
    End If
    
End Function


Private Sub tmrBlink_Timer()
    
    If tmrBlink.Enabled = True And tmrBlink.Interval = 4000 Then
        
        eye.Visible = False
        Image2.Visible = True
        tmrBlink.Enabled = False
        tmrClosed.Enabled = True
        
    End If
    
End Sub

Private Sub tmrClosed_Timer()
    
    If tmrClosed.Enabled = True And tmrClosed.Interval = 100 Then
        
        eye.Visible = True
        Image2.Visible = False
        tmrClosed.Enabled = False
        tmrBlink.Enabled = True
        
    End If
    
End Sub


