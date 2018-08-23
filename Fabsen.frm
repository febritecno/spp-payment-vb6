VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Fabsen 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entry Absen"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   ControlBox      =   0   'False
   Icon            =   "Fabsen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5640
      Top             =   3960
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
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   16
      ToolTipText     =   "mengikuti jam sistem"
      Top             =   330
      Width           =   1335
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "mengikuti tanggal sistem"
      Top             =   330
      Width           =   1335
   End
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "&Keluar"
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Fabsen.frx":0442
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "Fabsen.frx":0456
      TabIndex        =   7
      Top             =   2880
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   5175
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
         MaxLength       =   75
         TabIndex        =   18
         Top             =   1320
         Width           =   975
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
         MaxLength       =   75
         TabIndex        =   17
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   300
         Left            =   3120
         TabIndex        =   12
         Top             =   240
         Width           =   375
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
         Top             =   600
         Width           =   3255
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
         Width           =   1335
      End
      Begin VB.Data Data1 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Siswa"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NIS"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "Klik pada Kolom Nis untuk memilih data yang dihapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   5175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Daftar Absen "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   5175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Jam Absen"
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Absen"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Fabsen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ada As Boolean
'untuk menonaktifkan tombol command
Private Sub tomboltidakaktif()
cmdhapus.Enabled = False
End Sub
'Untuk mengaktikan tombol command
Private Sub tombolaktif()
cmdhapus.Enabled = True
End Sub

Private Sub Command1_Click()
frmcarisiswa.Show 1
End Sub

Private Sub DBGrid1_Click()
cmdhapus.Enabled = True
End Sub

Private Sub Text1_Change()
With Data1.Recordset
      .Index = "nisx"
      .Seek "=", Text1.Text
      If Not .NoMatch Then
         tampilkan
         textboxtdkaktif
         tombolaktif
         cmdsimpan.Enabled = True
      Else
         kosong
         cmdsimpan.Enabled = False
      End If
   End With
End Sub

'perintah penekanan tombol enter
Private Sub text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text5.SetFocus
End If
End Sub
Private Sub text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text6.SetFocus
End If
End Sub
Private Sub text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdsimpan.SetFocus
End If
End Sub
'untuk mengaktifkan textbox
Private Sub textboxaktif()
Text2.Enabled = True: Text5.Enabled = True: Text6.Enabled = True
End Sub
'untuk menonaktifkan textbox
Private Sub textboxtdkaktif()
Text2.Enabled = False: Text5.Enabled = False: Text6.Enabled = False
End Sub
'perintah pada saat tombol batal diklik
Private Sub cmdbatal_Click()
tomboltidakaktif
textboxtdkaktif
kosong
Text1.Enabled = True
Text1.SetFocus
End Sub
Private Sub cmdhapus_Click()
cmdhapus.Enabled = False
Data3.RecordSource = "Select * From tblabsen Where nis = '" & DBGrid1.Text & "' and cdate(tglabsen) = '" & Text3.Text & "'"
Data3.Refresh
With Data3.Recordset
     If Not .EOF Then
         Pesan = MsgBox("Yakin data ini akan dihapus ...?", vbYesNo, "Konfirmasi")
         If Pesan = vbYes Then
           .Delete
           Data3.Refresh
           Data2.Refresh
           cmdbatal_Click
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
Unload Me  'keluar dari form
End Sub
'perintah pada saat tombol simpan diklik
Private Sub cmdsimpan_Click()

If (Text1.Text = Empty) Then
      MsgBox "Data belum lengkap !", , "Konfirmasi"
      Exit Sub
Else ' jika sudah lengkap
   
   Data3.RecordSource = "Select * From tblabsen Where nis = '" & Text1.Text & "' and cdate(tglabsen) = '" & Text3.Text & "'"
   Data3.Refresh
   With Data3.Recordset
      If .EOF Then
           .AddNew
           !nis = Text1.Text
           !tglabsen = Text3.Text
           !jam = Text4.Text
           .Update
      Else
          MsgBox "Siswa tersebut sudah absen pada tanggal hari ini", 0 + 16, "Info"
          kosong
          Exit Sub
      End If
    End With
End If
Data3.Refresh
Data2.Refresh
kosong
cmdbatal_Click
End Sub
'untuk mengosongkan kembali textbox
Private Sub kosong()
Text5.Text = "": Text1.Text = "": Text2.Text = "": Text6.Text = ""
End Sub

'perintah pada saat form dijalankan
Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\DBADMINISTRASISEKOLAH.mdb"
Data2.DatabaseName = App.Path & "\DBADMINISTRASISEKOLAH.mdb"
Data3.DatabaseName = App.Path & "\DBADMINISTRASISEKOLAH.mdb"
Data1.RecordsetType = 0
Data2.RecordsetType = 0
Data3.RecordsetType = 1
Data1.RecordSource = "tblsiswa"
Data2.RecordSource = "tblabsen"
'Data2.RecordSource = "Select tblabsen.nis,tblsiswa.nama,tblsiswa.jk From tblabsen,tblsiswa Where tblabsen.nis=tblsiswa.nis and cdate(tblabsen.tglabsen) = '" & Text3.Text & "' order by tblabsen.nis"
Data2.Refresh
tomboltidakaktif
textboxtdkaktif
ada = True
Text3.Text = Date
Text4.Text = Time()
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
         cmdsimpan.Enabled = True
      Else
         MsgBox "Data nis tersebut tidak ada", 0 + 16, "Konfirmasi"
         kosong
         Exit Sub
      End If
   End With
End If
End Sub
'perintah untuk menampilkan data pada form textbox
Private Sub tampilkan()
With Data1.Recordset
  Text2.Text = !nama
  Text5.Text = !jk
  Text6.Text = !kelas
End With
End Sub
'Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = vbKeyReturn Then
'     SendKeys vbTab
'  End If
  
'End Sub
Private Sub Timer1_Timer()
Text4.Text = Time()
End Sub
