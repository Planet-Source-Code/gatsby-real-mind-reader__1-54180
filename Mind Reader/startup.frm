VERSION 5.00
Begin VB.Form startup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mind Reader"
         BeginProperty Font 
            Name            =   "Pepita MT"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   0
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "By Benjamin Blasen"
         BeginProperty Font 
            Name            =   "Pepita MT"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Shape box 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         DrawMode        =   15  'Merge Pen Not
         FillStyle       =   6  'Cross
         Height          =   5295
         Left            =   -5160
         Shape           =   2  'Oval
         Top             =   -360
         Width           =   6735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5160
      Top             =   3600
   End
End
Attribute VB_Name = "startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    box.Left = -10000
    
End Sub

Private Sub Timer1_Timer()
    
    If box.Left >= startup.Width Then
    
        Unload Me
        mindReader.Show
        
    Else
    
        box.Left = box.Left + 20
    
    End If
    
End Sub
