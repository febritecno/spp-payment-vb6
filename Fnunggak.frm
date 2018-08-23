VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Siswa Menunggak"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   960
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\Source Code - Tugas-KP\Tugas-KP\madrasah.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblsiswa"
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Baru"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Fnunggak.frx":0000
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "Fnunggak.frx":0014
      TabIndex        =   14
      Top             =   4320
      Width           =   8055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Siswa"
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
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
         Left            =   1560
         TabIndex        =   21
         Top             =   3240
         Width           =   2535
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
         Left            =   1560
         TabIndex        =   13
         Top             =   2760
         Width           =   2535
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
         Left            =   1560
         TabIndex        =   12
         Top             =   2280
         Width           =   1215
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
         Left            =   1560
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
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
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
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
         Left            =   1560
         TabIndex        =   5
         Top             =   840
         Width           =   4455
      End
      Begin VB.CommandButton Tombol1 
         Caption         =   "..."
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
         Left            =   4200
         TabIndex        =   3
         Top             =   360
         Width           =   375
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
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   20
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Biaya"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   11
         Top             =   3240
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   10
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SPP Bulan"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   8
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nis"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   225
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ada As Boolean

Private Sub cmdBatal_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo1.Text = ""
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Combo1.Enabled = False
Tombol1.Enabled = False
cmdSimpan.Enabled = False
cmdHapus.Enabled = False
End Sub

Private Sub cmdHapus_Click()
cmdHapus.Enabled = False
Data1.RecordSource = "Select * From tblnunggak Where nis = '" & DBGrid1.Text & "'"
Data1.Refresh
With Data1.Recordset
     If Not .EOF Then
         Pesan = MsgBox("Yakin data ini akan dihapus ...?", vbYesNo, "Konfirmasi")
         If Pesan = vbYes Then
           .Delete
           Data1.Refresh
           'Data2.Refresh
           cmdBatal_Click
           cmdSimpan.Enabled = True
         Else
           Exit Sub
         End If
      End If
End With '
End Sub

Private Sub cmdSimpan_Click()
If (Text1.Text = Empty Or Combo1.Text = Empty Or Text2.Text = Empty Or Text3.Text = Empty Or Text4.Text = Empty Or Text5.Text = Empty) Then
      MsgBox "Data belum lengkap !", , "Konfirmasi"
      Exit Sub
Else ' jika sudah lengkap
   
   Data1.RecordSource = "Select * From tblnunggak Where nis = '" & Text1.Text & "' and sppbulan = '" & Combo1.Text & "'"
   Data1.Refresh
   With Data1.Recordset
      If .EOF Then
           .AddNew
           !sppbulan = Combo1.Text
           !tahun = Text4.Text
           !nis = Text1.Text
           !nama = Text2.Text
           !kelas = Text3.Text
           !biaya = Text6.Text
           .Update
           'Text8.Text = ""
           'Text9.Text = ""
      Else
          MsgBox "Ada kesalahan dalam pemasukan data", 0 + 16, "Info"
          'kosong
          Text1.Text = ""
          Text2.Text = ""
          Text3.Text = ""
          Text4.Text = ""
          Text5.Text = ""
          Text6.Text = ""
          Combo1.Text = ""
          Text1.Enabled = False
          Text2.Enabled = False
          Text3.Enabled = False
          Text4.Enabled = False
          Text5.Enabled = False
          Combo1.Enabled = False
          Text6.Enabled = False
          Tombol1.Enabled = False
          cmdSimpan.Enabled = False
          cmdHapus.Enabled = False
          cmdBatal.Enabled = False
          Exit Sub
      End If
    End With
End If
Data1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Combo1.Enabled = False
Text6.Enabled = False
Tombol1.Enabled = False
cmdSimpan.Enabled = False
cmdHapus.Enabled = False
cmdBatal.Enabled = False
'Data2.Refresh
'kosong
'cmdbatal_Click
End Sub

Private Sub Command1_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Combo1.Enabled = True
Text6.Enabled = True
Tombol1.Enabled = True
cmdSimpan.Enabled = True
cmdBatal.Enabled = True
End Sub
Private Sub tampilkan()
With Data2.Recordset
  Text2.Text = !nama
  Text3.Text = !kelas
  Text5.Text = !Status
  Text6.Text = !spp
  'Combo1.Text = !JnsKelamin
  'DTPicker1.Value = !TglLahir
  'Text3.Text = !Alamat
  'Text4.Text = !Jabatan
End With
End Sub

Private Sub Command3_Click()

End Sub

Private Sub DBGrid1_Click()
cmdHapus.Enabled = True
End Sub

Private Sub Text1_Change()
With Data2.Recordset
      .Index = "nisx"
      .Seek "=", Text1.Text
      If Not .NoMatch Then
         tampilkan
        Else
        Text2.Text = ""
        Text3.Text = ""
         End If
   End With
End Sub



Private Sub Tombol1_Click()
frmcarisiswa.Show
End Sub

Private Sub Command2_Click()
Fmenuutama.Show
Fmenuutama.menuaktif
Unload Me
ada = False
End Sub

Private Sub Form_Load()
ada = True
Data1.DatabaseName = App.Path & "\madrasah.mdb"
Data1.RecordsetType = 1
Data1.RecordSource = "tblnunggak"
Data2.DatabaseName = App.Path & "\madrasah.mdb"
Data2.RecordsetType = 0
Data2.RecordSource = "tblsiswa"
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Combo1.Enabled = False
Text6.Enabled = False
Tombol1.Enabled = False
cmdSimpan.Enabled = False
cmdHapus.Enabled = False
cmdBatal.Enabled = False
Combo1.AddItem "Januari"
Combo1.AddItem "Februari"
Combo1.AddItem "Maret"
Combo1.AddItem "April"
Combo1.AddItem "Mei"
Combo1.AddItem "Juni"
Combo1.AddItem "Juli"
Combo1.AddItem "Agustus"
Combo1.AddItem "September"
Combo1.AddItem "Oktober"
Combo1.AddItem "Nopember"
Combo1.AddItem "Desember"
End Sub


