VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "View Users"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Users"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2415
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   1
      Appearance      =   0
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "login"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New User"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmAddNew.Show
End Sub

Private Sub Command2_Click()
    Form1.Width = 7860
    Command2.Visible = False
    Command3.Visible = True
End Sub
Private Sub Command3_Click()
    Form1.Width = 3825
    Command2.Visible = True
    Command3.Visible = False
End Sub
Private Sub Form_Load()
    Data1.DatabaseName = (App.Path & "\testlogin.mdb")
    Data1.RecordSource = "login"
End Sub
