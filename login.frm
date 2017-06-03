VERSION 5.00
Begin VB.Form flogin 
   Caption         =   "Login"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton batal 
      Caption         =   "Batal"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton login 
      Caption         =   "Login"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox password 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox username 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "flogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub batal_Click()
username.Text = ""
password.Text = ""
batal.Enabled = False
login.Enabled = False
password.Enabled = False
End Sub

Private Sub Form_Load()
username.MaxLength = 15
password.MaxLength = 15
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
fberanda.Show
End Sub

Private Sub login_Click()
masuk
End Sub

Sub masuk()
BukaDataBase
If username.Text = "" Or password.Text = "" Then
MsgBox "Masukkan data login"
Else
rsLogin.Open "SELECT * FROM user WHERE username='" & username.Text & "' AND password='" & password.Text & "'", KonekDb, adOpenDynamic, adLockBatchOptimistic
    If rsLogin.EOF Then
    MsgBox "Password atau username salah!", vbOKOnly, "Peringatan"
    username.Text = ""
    password.Text = ""
    password.Enabled = False
    username.SetFocus
    batal.Enabled = False
    login.Enabled = False
    KonekDb.Close
    Else
    Unload Me
    MsgBox "Selamat datang " & rsLogin!nama
    fberanda.Show
    fberanda.master.Enabled = True
    fberanda.transaksi.Enabled = True
    fberanda.fileLogout.Enabled = True
    fberanda.fileLogin.Enabled = False
    fberanda.statusbar.Panels(2) = "User : " & rsLogin!nama
    End If
End If
End Sub
Private Sub password_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then masuk
End Sub

Private Sub username_Change()
password.Enabled = True
batal.Enabled = True
login.Enabled = True
End Sub
