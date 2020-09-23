VERSION 5.00
Begin VB.Form frmAddNew 
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Add"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "login"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSubmit_Click()
   
    Dim samepassword As Boolean
    
    If txtPassword.Text = txtConfirmPassword.Text Then
        samepassword = True
    Else
        samepassword = False
        MsgBox "The passwords do not match", , "Login"
        txtConfirmPassword.SetFocus
        Exit Sub
    End If
    With Data1.Recordset
    .MoveLast
    .AddNew
    .Fields!UserName = txtUserName.Text
    .Fields!Password = txtPassword.Text
    .Update
    MsgBox "Username and password added"
    End With
    txtUserName.Text = ""
    txtPassword.Text = ""
    txtConfirmPassword.Text = ""
    txtUserName.SetFocus
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Data1.DatabaseName = (App.Path & "\testlogin.mdb")
    Data1.RecordSource = "login"
End Sub
