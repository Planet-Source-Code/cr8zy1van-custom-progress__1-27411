VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Progress"
   ClientHeight    =   1500
   ClientLeft      =   11385
   ClientTop       =   780
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3480
   Begin Project1.Meter3 proControl4 
      Height          =   525
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   926
      Border          =   0
   End
   Begin Project1.Meter3 proControl3 
      Height          =   525
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   926
      Border          =   0
   End
   Begin Project1.Meter3 proControl2 
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   926
      Border          =   0
   End
   Begin Project1.Meter3 proControl1 
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   926
      Border          =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   2640
   End
   Begin Project1.Meter3 proControl 
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   926
      Border          =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_DblClick()
Timer1.Enabled = False
Unload Me
End Sub


Private Sub Timer1_Timer()
    Randomize 'make sure it is totaly random
    random = Int((Rnd * 50) + 1) 'you can change the 100 To something Else
    For I = 0 To 9 'go through all the controls starting With 0
        'Me.Caption = (random * 2) & "%" 'make the label's caption the random number
    Next I 'go To the next label
    proControl.Value = random
    proControl1.Value = random / 2
    proControl2.Value = random / 2 * 1.5
    proControl3.Value = random - 5 / 3 + 2
    proControl4.Value = random / 5 * 1
    
End Sub
