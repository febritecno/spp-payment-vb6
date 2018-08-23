VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Fsiswa 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entry Siswa"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   ControlBox      =   0   'False
   Icon            =   "Fsiswa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "KELUAR"
      Height          =   975
      Left            =   6240
      TabIndex        =   35
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "BATAL"
      Height          =   975
      Left            =   6240
      TabIndex        =   34
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "HAPUS"
      Height          =   975
      Left            =   6240
      TabIndex        =   33
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "EDIT"
      Height          =   975
      Left            =   6240
      TabIndex        =   32
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "SIMPAN"
      Height          =   975
      Left            =   6240
      TabIndex        =   31
      Top             =   480
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Source Code - Tugas-KP\Tugas-KP\madrasah.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblsiswa"
      Top             =   8520
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Fsiswa.frx":0442
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "Fsiswa.frx":0456
      TabIndex        =   6
      Top             =   6480
      Width           =   7575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   6135
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   6015
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142999553
         CurrentDate     =   43211
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3840
         TabIndex        =   28
         Top             =   5040
         Width           =   2055
      End
      Begin VB.ComboBox cmbStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1080
         TabIndex        =   26
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   4560
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   3600
         Width           =   4095
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin MSComCtl2.UpDown Up1 
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   5640
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Value           =   2009
         Max             =   9999
         Min             =   2009
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtthn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   5640
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   300
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         MaxLength       =   75
         TabIndex        =   1
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Biaya SPP"
         Height          =   195
         Index           =   8
         Left            =   3000
         TabIndex        =   27
         Top             =   5040
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   25
         Top             =   5040
         Width           =   450
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Ibu"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   23
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Ayah"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   21
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Siswa"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   19
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Agama"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Lahir"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Ajaran"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Siswa"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NIS"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Fsiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ada As Boolean
'untuk menonaktifkan tombol command
Private Sub tomboltidakaktif()
cmdedit.Enabled = False: cmdhapus.Enabled = False
End Sub
'Untuk mengaktikan tombol command
Private Sub tombolaktif()
cmdedit.Enabled = True: cmdhapus.Enabled = True
End Sub

Private Sub cmbStatus_Click()
If cmbStatus.Text = "Kurang Mampu" Then
Text8.Text = "70000"
ElseIf cmbStatus.Text = "Cukup Mampu" Then
Text8.Text = "90000"
Else
Text8.Text = "150000"
End If
End Sub

Private Sub Command1_Click()
frmcarisiswa.Show 1
End Sub

Private Sub Text1_Change()
With Data1.Recordset
      .Index = "nisx"
      .Seek "=", Text1.Text
      If Not .NoMatch Then
         tampilkan
         textboxtdkaktif
         tombolaktif
         cmdsimpan.Enabled = False
      Else
         kosong
         textboxaktif
      End If
   End With
End Sub

'perintah penekanan tombol enter
Private Sub text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If
End Sub
Private Sub combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If
End Sub
Private Sub combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdsimpan.SetFocus
End If
End Sub
'untuk mengaktifkan textbox
Private Sub textboxaktif()
Text2.Enabled = True: Combo1.Enabled = True: Combo2.Enabled = True: Text3.Enabled = True: DTPicker1.Enabled = True: Text4.Enabled = True: Text5.Enabled = True: Text6.Enabled = True: Text7.Enabled = True: Text8.Enabled = True: cmbStatus.Enabled = True
End Sub
'untuk menonaktifkan textbox
Private Sub textboxtdkaktif()
Text2.Enabled = False: Combo1.Enabled = False: Combo2.Enabled = False: Text3.Enabled = False: DTPicker1.Enabled = False: Text4.Enabled = False: Text5.Enabled = False: Text6.Enabled = False: Text7.Enabled = False: Text8.Enabled = False: cmbStatus.Enabled = False
End Sub
'perintah pada saat tombol batal diklik
Private Sub cmdBatal_Click()
tomboltidakaktif
textboxtdkaktif
kosong
Text1.Enabled = True
Text1.Text = ""
Text1.SetFocus
End Sub
'perintah pada saat tombol edit diklik
Private Sub cmdedit_Click()
cmdsimpan.Enabled = True
textboxaktif
cmdhapus.Enabled = False
cmdedit.Enabled = False
Text1.Enabled = False
Text2.SetFocus
End Sub
'perintah pada saat tombol hapus diklik
Private Sub cmdHapus_Click()
With Data1.Recordset
      .Index = "nisx"
      .Seek "=", Text1.Text
      If Not .NoMatch Then
         Pesan = MsgBox("Yakin data ini akan dihapus ...?", vbYesNo, "Konfirmasi")
         If Pesan = vbYes Then
           .Delete
           Data1.Refresh
           cmdBatal_Click
           cmdsimpan.Enabled = True
         Else
           Exit Sub
         End If
      End If
End With '
End Sub
'perintah pada saat tombol keluar diklik
Private Sub cmdkeluar_Click()
ada = False
Fmenuutama.menuaktif
Fmenuutama.Show
Unload Me  'keluar dari form
End Sub
'perintah pada saat tombol simpan diklik
Private Sub cmdSimpan_Click()

If (Text1.Text = Empty Or Text2.Text = Empty Or Combo1.Text = Empty Or Combo2.Text = Empty) Then
      MsgBox "Data belum lengkap !", , "Konfirmasi"
      Exit Sub
Else ' jika sudah lengkap
   With Data1.Recordset
      .Index = "nisx"
      .Seek "=", Text1.Text
      If .NoMatch Then
           .AddNew
           !nis = Text1.Text
           !nama = Text2.Text
           !kelas = Combo2.Text
           !tmptlahir = Text3.Text
           !tgllahir = DTPicker1.Value
           !jnskelamin = Combo1.Text
           !agama = Text4.Text
           !alamat = Text5.Text
           !nmayah = Text6.Text
           !nmibu = Text7.Text
           !ta = txtthn.Text
           !Status = cmbStatus.Text
           !spp = Text8.Text
           .Update
      Else
           .Edit
           !nama = Text2.Text
           !kelas = Combo2.Text
           !tmptlahir = Text3.Text
           !tgllahir = DTPicker1.Value
           !jnskelamin = Combo1.Text
           !agama = Text4.Text
           !alamat = Text5.Text
           !nmayah = Text6.Text
           !nmibu = Text7.Text
           !ta = txtthn.Text
           !Status = cmbStatus.Text
           !spp = Text8.Text
           .Update
      End If
    End With
End If
Data1.Refresh
kosong
cmdBatal_Click
End Sub
'untuk mengosongkan kembali textbox
Private Sub kosong()
Combo1.Text = "": Text2.Text = "": Combo2.Text = ""
End Sub

'perintah pada saat form dijalankan
Private Sub Form_Load()
On Error Resume Next
DTPicker1 = Format(Date, "dd MM yyyy")
Data1.DatabaseName = App.Path & "\madrasah.mdb"
Data1.RecordsetType = 0
Data1.RecordSource = "tblsiswa"
tomboltidakaktif
textboxtdkaktif
ada = True
Combo1.AddItem "Laki-Laki"
Combo1.AddItem "Perempuan"
Combo2.AddItem "I"
Combo2.AddItem "II"
Combo2.AddItem "III"
Combo2.AddItem "IV"
Combo2.AddItem "V"
Combo2.AddItem "VI"
Combo2.AddItem "VII"
Combo2.AddItem "VIII"
Combo2.AddItem "IX"
Combo2.AddItem "X"
Combo2.AddItem "XI"
Combo2.AddItem "XII"
cmbStatus.AddItem "Kurang Mampu"
cmbStatus.AddItem "Cukup Mampu"
cmbStatus.AddItem "Sangat Mampu"
Up1.Value = Year(Date)
txtthn.Text = Trim(Str(Up1.Value)) + "/" + Trim(Str(Up1.Value + 1))
End Sub
'perintah pada saat penekanan tomblol enter pada textbox
Private Sub text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   With Data1.Recordset
      .Index = "nisx"
      .Seek "=", Text1.Text
      If Not .NoMatch Then
         tampilkan
         textboxtdkaktif
         tombolaktif
         cmdsimpan.Enabled = False
      Else
         textboxaktif
         Text2.SetFocus
      End If
   End With
End If
End Sub
'perintah untuk menampilkan data pada form textbox
Private Sub tampilkan()
With Data1.Recordset
  Text2.Text = !nama
  Combo1.Text = !jnskelamin
  Combo2.Text = !kelas
  txtthn.Text = !ta
  Text3.Text = !tmptlahir
  DTPicker1.Value = !tgllahir
  Text4.Text = !agama
  Text5.Text = !alamat
  Text6.Text = !nmayah
  Text7.Text = !nmibu
  cmbStatus.Text = !Status
  Text8.Text = !spp
End With
End Sub
'Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = vbKeyReturn Then
'     SendKeys vbTab
'  End If
  
'End Sub
Private Sub Up1_Change()
txtthn.Text = Trim(Str(Up1.Value)) + "/" + Trim(Str(Up1.Value + 1))
End Sub
