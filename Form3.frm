VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "            "
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810
   LinkTopic       =   "Form3"
   ScaleHeight     =   6360
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text6 
      DataField       =   "cantidad"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   21
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "foto"
      Height          =   495
      Left            =   8400
      TabIndex        =   19
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "eliminar"
      Height          =   495
      Left            =   7320
      TabIndex        =   18
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "modificar"
      Height          =   495
      Left            =   6240
      TabIndex        =   17
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "nuevo"
      Height          =   495
      Left            =   5160
      TabIndex        =   16
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "guardar"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   13
      Top             =   3000
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   2655
      Left            =   4080
      ScaleHeight     =   2595
      ScaleWidth      =   2595
      TabIndex        =   12
      Top             =   1560
      Width           =   2655
      Begin VB.Image Image3 
         DataSource      =   "Adodc1"
         Height          =   2415
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   5400
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
      RecordSource    =   "farmacia"
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
   Begin VB.TextBox Text5 
      DataField       =   "precio"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      DataField       =   "tipo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataField       =   "cantidad de producto"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "codigo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "cantidad"
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
      TabIndex        =   20
      Top             =   4200
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   2295
      Left            =   4320
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label7 
      DataField       =   "foto"
      DataSource      =   "Adodc1"
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
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "precio"
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
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "tipo"
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
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "cantidad de producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "nombre"
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
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "codigo"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Farmacia"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   6255
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Form3"
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
    Image2.Picture = LoadPicture(x & "\" & Label8.Caption)
End Sub

Private Sub Command2_Click()
    Adodc1.Recordset.MoveNext
    
    If Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveFirst
    End If
    x = App.Path
    Image2.Picture = LoadPicture(x & "\" & Label8.Caption)
End Sub
Private Sub Command3_Click()
    FileCopy CommonDialog1.FileName, App.Path & "\\" & CommonDialog1.FileTitle
     Adodc1.Recordset.Update
    Adodc1.Recordset.MoveFirst
    x = App.Path
    Image2.Picture = LoadPicture(x & "\" & Label8.Caption)

    
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
    Command7.Enabled = True
    
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
    Image2.Picture = LoadPicture(Label8.Caption)
    
End Sub

Private Sub Command5_Click()
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveFirst
    x = App.Path
    Image2.Picture = LoadPicture(x & "\" & Label8.Caption)
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
    Command7.Enabled = True
End Sub

Private Sub Command7_Click()
    CommonDialog1.ShowOpen
    Image2.Picture = LoadPicture(CommonDialog1.FileName)
    Label8.Caption = CommonDialog1.FileTitle
    
    If Label8.Caption = "" Then
        MsgBox ("seleccione una imagen")
    Else
         Label8.Caption = CommonDialog1.FileTitle
    End If
End Sub

Private Sub Form_Load()
    x = App.Path
    Image2.Picture = LoadPicture(x & "\" & Label8.Caption)
    
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
    Command6.Enabled = True
    Command7.Enabled = False
    

End Sub



