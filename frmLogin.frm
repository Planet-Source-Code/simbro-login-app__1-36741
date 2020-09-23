VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2850
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1683.874
   ScaleMode       =   0  'User
   ScaleWidth      =   3675.973
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "login"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim db As Database
    Dim rs As Recordset
    
    Set db = OpenDatabase(App.Path & "\testlogin.mdb")
    Set rs = db.OpenRecordset("login")
    
    Do While Not rs.EOF
        If rs.Fields("username") = (txtUserName.Text) And _
        rs.Fields("password") = (txtPassword.Text) Then
        Form1.Show
        Unload Me
        Exit Sub
    Else
        rs.MoveNext
        End If
    Loop
    txtPassword.Text = ""
    MsgBox "Incorrect Password!", vbCritical
End Sub

Private Sub frmLogin_Load()
    Data1.DatabaseName = (App.Path & "\testlogin.mdb")
    Data1.RecordSource = "login"
End Sub
