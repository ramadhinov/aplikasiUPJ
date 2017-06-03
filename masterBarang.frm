VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fmasterBarang 
   Caption         =   "Master Barang"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9585
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid tabeldata 
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4260
      _Version        =   393216
   End
   Begin VB.TextBox cari 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Text            =   "Cari barang"
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      Caption         =   " Menu "
      Height          =   3975
      Left            =   7920
      TabIndex        =   15
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton hapus 
         Caption         =   "&Hapus"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton batal 
         Caption         =   "&Batal"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton baru 
         Caption         =   "&Baru"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Data Barang "
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7695
      Begin VB.ComboBox kategori 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   5
         Text            =   "-- Pilih Kategori --"
         Top             =   2880
         Width           =   4935
      End
      Begin VB.TextBox stok_awal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox spesifikasi 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   2520
         TabIndex        =   4
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox nama 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox id 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Stok Masuk"
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
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Kategori"
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
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Spesifikasi"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Barang"
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
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "ID Barang"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "fmasterBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub baru_Click()
If baru.Caption = "&Baru" Then
 baru.Caption = "&Simpan"
 enable
 hapus.Enabled = False
ElseIf baru.Caption = "&Simpan" Then
 baru.Caption = "&Baru"
ElseIf baru.Caption = "&Edit" Then
 baru.Caption = "&Update"
ElseIf baru.Caption = "&Update" Then
 baru.Caption = "&Baru"
End If
 
End Sub

Private Sub batal_Click()
bersih
disable
baru.Caption = "&Baru"
End Sub

Private Sub cari_Click()
cari.Text = ""
End Sub

Sub bersih()
id.Text = ""
nama.Text = ""
spesifikasi.Text = ""
kategori.Text = "-- Pilih Kategori --"
stok_awal.Text = ""
cari.Text = "Cari barang"
End Sub

Sub disable()
nama.Enabled = False
spesifikasi.Enabled = False
kategori.Enabled = False
stok_awal.Enabled = False
batal.Enabled = False
hapus.Enabled = False
End Sub

Sub enable()
nama.Enabled = True
spesifikasi.Enabled = True
kategori.Enabled = True
stok_awal.Enabled = True
batal.Enabled = True
hapus.Enabled = True
End Sub

Private Sub cari_GotFocus()
cari.Text = ""
End Sub

Private Sub cari_LostFocus()
cari.Text = "Cari barang"
End Sub

Private Sub Form_Load()
kategori.AddItem ("Keyboard dan Mouse")
End Sub
