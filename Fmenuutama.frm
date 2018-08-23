VERSION 5.00
Begin VB.Form Fmenuutama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEMBAYARAN SPP"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Fmenuutama.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "EXIT"
      Height          =   855
      Left            =   9720
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LAPORAN"
      Height          =   855
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BAYAR SPP"
      Height          =   855
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DATA SISWA"
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Fmenuutama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
menutdkaktif
End Sub
Sub menutdkaktif()
Command2.Enabled = False: Command3.Enabled = False: Command4.Enabled = False
End Sub
Sub menuaktif()
Command2.Enabled = True: Command3.Enabled = True: Command4.Enabled = True
End Sub



Private Sub Timer1_Timer()
Label3.Caption = "JAM : " + Format(Time())
Label4.Caption = "TANGGAL : " + Format(Date)
End Sub

Private Sub Command1_Click()
menutdkaktif
Flogin.Show
Unload Me
End Sub

Private Sub Command2_Click()
Fsiswa.Show
Unload Me
End Sub

Private Sub Command3_Click()
FBayarSPP.Show
Unload Me
End Sub

Private Sub command4_Click()
FLaporan.Show
Unload Me
End Sub

Private Sub command5_Click()
End
End Sub

