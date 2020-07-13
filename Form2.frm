VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\hospital\Database1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\hospital\Database1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "personalmedico"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "guardar"
      Height          =   495
      Left            =   4440
      TabIndex        =   21
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   " foto"
      Height          =   495
      Left            =   8640
      TabIndex        =   20
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "eliminar"
      Height          =   495
      Left            =   7560
      TabIndex        =   19
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "modificar"
      Height          =   495
      Left            =   6480
      TabIndex        =   18
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "nuevo"
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   495
      Left            =   8640
      TabIndex        =   16
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   495
      Left            =   7800
      TabIndex        =   15
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0FF&
      FillColor       =   &H00FFC0FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   4920
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   14
      Top             =   1440
      Width           =   2655
      Begin VB.Image image2 
         DataSource      =   "Adodc1"
         Height          =   2175
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox Text6 
      DataField       =   "fecha de inicio"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      DataField       =   "sueldo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataField       =   "especialidad"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "cargo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "CUI"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label8 
      DataField       =   "foto"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Facha de inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Sueldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Especialidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "personal medico"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   5940
      Left            =   0
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9825
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Adodc1.Recordset.MoveLast
    
    If Adodc1.Recordset.BOF Then
        adocd1.Recordset.MoveLast
    End If
     x = App.Path
    image2.Picture = LoadPicture(x & "\" & Label8.Caption)
End Sub

Private Sub Command2_Click()
    Adodc1.Recordset.MoveNext
   
    If Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveFirst
    End If
      x = App.Path
    image2.Picture = LoadPicture(x & "\" & Label8.Caption)
    
End Sub
Private Sub Command3_Click()
    FileCopy CommonDialog1.FileName, App.Path & "\\" & CommonDialog1.FileTitle
     Adodc1.Recordset.Update
    Adodc1.Recordset.MoveFirst
    x = App.Path
    image2.Picture = LoadPicture(x & "\" & Label8.Caption)

    
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Command3.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = False
    
End Sub

Private Sub Command4_Click()
    Adodc1.Recordset.AddNew
    
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Command3.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = True
    Text1.SetFocus
    
    Label11.Caption = ""
    image2.Picture = LoadPicture(Label8.Caption)
    
End Sub

Private Sub Command5_Click()
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveFirst
    x = App.Path
    image2.Picture = LoadPicture(x & "\" & Label8.Caption)
End Sub

Private Sub Command6_Click()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Command3.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
End Sub

Private Sub Command7_Click()
    CommonDialog1.ShowOpen
    image2.Picture = LoadPicture(CommonDialog1.FileName)
    Label8.Caption = CommonDialog1.FileTitle
    
    If Label8.Caption = "" Then
        MsgBox ("seleccione una imagen")
    Else
         Label8.Caption = CommonDialog1.FileTitle
    End If
End Sub

Private Sub Form_Load()
     x = App.Path
    image2.Picture = LoadPicture(x & "\" & Label8.Caption)
    
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = False
    Command7.Enabled = False
    

End Sub


