VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm fberanda 
   BackColor       =   &H8000000C&
   Caption         =   "Unit Produksi dan Jasa"
   ClientHeight    =   7845
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9585
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar statusbar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7590
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2470
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3774
            MinWidth        =   3774
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu fileLogin 
         Caption         =   "Login"
      End
      Begin VB.Menu fileCetak 
         Caption         =   "Cetak"
      End
      Begin VB.Menu fileLogout 
         Caption         =   "Logout"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu master 
      Caption         =   "Master"
      Enabled         =   0   'False
      Begin VB.Menu masterPengguna 
         Caption         =   "Pengguna"
      End
      Begin VB.Menu MasterBarang 
         Caption         =   "Barang"
      End
   End
   Begin VB.Menu transaksi 
      Caption         =   "Transaksi"
      Enabled         =   0   'False
      Begin VB.Menu transaksiPembayaran 
         Caption         =   "Pembayaran"
      End
      Begin VB.Menu transaksiStok 
         Caption         =   "Penambahan Stok"
      End
   End
   Begin VB.Menu tentang 
      Caption         =   "Tentang"
   End
End
Attribute VB_Name = "fberanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fileLogin_Click()
Unload Me
flogin.Show
End Sub

Private Sub MasterBarang_Click()
keluar
fmasterBarang.Show
End Sub

Private Sub masterPengguna_Click()
keluar
fmasterPengguna.Show
End Sub

Private Sub fileLogout_Click()
Dim x As String
x = MsgBox("Apakah anda yakin keluar?", vbYesNo + vbExclamation, "Peringatan!")
    If x = vbYes Then keluarin
End Sub

Sub keluarin()
Unload Me
fberanda.fileLogout.Enabled = False
fberanda.master.Enabled = False
fberanda.transaksi.Enabled = False
fberanda.fileLogin.Enabled = True
fberanda.Show
KonekDb.Close
KonekDb.Open
End Sub

Private Sub MDIForm_Load()
statusbar.Panels(1) = Format(Now, "DD-MM-YYYY")
statusbar.Panels(2) = "Status : Belum Login"
End Sub

Private Sub tentang_Click()
keluar
ftentang.Show
End Sub

Sub keluar()
Unload fmasterBarang
Unload fmasterPengguna
End Sub
