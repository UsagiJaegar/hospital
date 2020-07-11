VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0FF&
      Cancel          =   -1  'True
      Caption         =   "pacientes"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "famacia"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "personal medico"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hospital"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Visible = True
    Form1.Visible = False
    
End Sub

Private Sub Command2_Click()
    Form3.Visible = True
    Form1.Visible = False
End Sub

Private Sub Command3_Click()
    Form4.Visible = True
    Form1.Visible = False
    
End Sub
