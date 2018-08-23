VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FLaporan 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crtCetak 
      Left            =   11400
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox TxtNis 
      Height          =   375
      Left            =   8640
      TabIndex        =   28
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "D:\Source Code - Tugas-KP\Tugas-KP\madrasah.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblsiswa"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Siswa Yang Membayar"
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   600
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   975
      Left            =   600
      TabIndex        =   5
      Top             =   600
      Width           =   5895
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Bayar"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "s.d."
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H00C0C000&
      Caption         =   "Siswa Yang Menunggak"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   600
      TabIndex        =   10
      Top             =   1920
      Width           =   5895
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   975
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
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SPP Bulan"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Laporan Seluruh Siswa"
      Height          =   3135
      Left            =   120
      TabIndex        =   27
      Top             =   240
      Width           =   6975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Laporan Pembayaran Per Siswa"
      Height          =   3135
      Left            =   7200
      TabIndex        =   18
      Top             =   240
      Width           =   5895
      Begin VB.CommandButton Command5 
         Caption         =   "Cetak Laporan Per Siswa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   2520
         Width           =   5415
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3360
         TabIndex        =   25
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   143196161
         CurrentDate     =   41636
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1800
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   143196161
         CurrentDate     =   41636
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
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
         Left            =   4920
         TabIndex        =   21
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox TxtNama 
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sampai Tanggal"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   23
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dari Tanggal"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   22
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Siswa"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Top             =   840
         Width           =   885
      End
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "FLaporan.frx":0000
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "FLaporan.frx":0014
      TabIndex        =   17
      Top             =   3720
      Width           =   12975
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   7200
      Width           =   8415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FLaporan.frx":09E7
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "FLaporan.frx":09FB
      TabIndex        =   3
      Top             =   3720
      Width           =   12975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Daftar Siswa Yang membayar "
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
      TabIndex        =   4
      Top             =   3480
      Width           =   12975
   End
End
Attribute VB_Name = "FLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ada As Boolean
Private Sub Command1_Click()
On Error Resume Next
If opt1 = True Then
    If Text1.Text = Empty Then
       Exit Sub
    End If
    If Text2.Text = Empty Then
       Exit Sub
    End If
    Label7.Caption = "Daftar Siswa Yang Membayar"
    DBGrid1.Visible = True
    DBGrid2.Visible = False
    'Data1.RecordSource = "Select * From tblbayarspp Where cdate(tglbayar) = '" & Text1.Text & "' order by nis"
    Data1.RecordSource = "Select tblbayarspp.nis,tblsiswa.nama,tblsiswa.kelas, tblbayarspp.sppbulan,tblbayarspp.tahun,tblbayarspp.tglbayar From tblbayarspp,tblsiswa Where tblbayarspp.nis=tblsiswa.nis and (cdate(tblbayarspp.tglbayar) >= '" & Format(Text1.Text, "dd/mm/yyyy") & "' and cdate(tblbayarspp.tglbayar) <= '" & Format(Text2.Text, "dd/mm/yyyy") & "') order by tblbayarspp.nis"
    Data1.Refresh
    With Data1.Recordset
          If Not .EOF Then
             Data2.RecordSource = "Select tblbayarspp.nis,tblsiswa.nama,tblsiswa.kelas, tblbayarspp.sppbulan,tblbayarspp.tahun,tblbayarspp.tglbayar From tblbayarspp,tblsiswa Where tblbayarspp.nis=tblsiswa.nis and (cdate(tblbayarspp.tglbayar) >= '" & Format(Text1.Text, "dd/mm/yyyy") & "' and cdate(tblbayarspp.tglbayar) <= '" & Format(Text2.Text, "dd/mm/yyyy") & "') order by tblbayarspp.nis"
             Data2.Refresh
             Command3.Enabled = True
          Else
             MsgBox "Tidak ada Yang membayar pada Tanggal tersebut ", 0 + 16, "Konfirmasi"
             Command3.Enabled = False
             Exit Sub
          End If
    End With
End If
If opt2 = True Then
    If Combo1.Text = Empty Then
       Exit Sub
    End If
    If Text3.Text = Empty Then
       Exit Sub
    End If
     Label7.Caption = "Daftar Siswa Yang Menunggak"
     DBGrid1.Visible = False
     DBGrid2.Visible = True
     Command3.Enabled = True
     Data3.RecordSource = "Select nis,nama,kelas,sppbulan,biaya From tblnunggak Where sppbulan like '*" & Combo1.Text & "*' and tahun like '*" & Text3.Text & "*' order by tblnunggak.nis"
     Data3.Refresh
    'With Data3.Recordset
    '      If .EOF Then
    '            Data3.RecordSource = "Select nis,nama,kelas,sppbulan,biaya From tblnunggak Where tblnunggak.sppbulan = '" & Combo1.Text & "' and tblnunggak.tahun = '" & Text3.Text & "' order by tblnunggak.nis"
    '            Data3.Refresh
    '            Command3.Enabled = True
          'Else
          '   DBGrid1.Refresh
          '   MsgBox "Tidak ada data yang diinput untuk siswa menunggak", 0 + 16, "Konfirmasi"
          '   Command3.Enabled = False
          '   Exit Sub
    '      End If
    'End With
End If

End Sub

Private Sub Command2_Click()
ada = False
Fmenuutama.menuaktif
Fmenuutama.Show
Unload Me
End Sub
Private Sub CetakSPP()
Me.MousePointer = 11
With crtCetak
        .Reset
        .ReportFileName = App.Path & "\LapBayarSPP.rpt"
        '.Password = Chr(10) & "amir"
        .DataFiles(0) = App.Path & "\madrasah.mdb"
        .SelectionFormula = "{tblbayarspp.tglbayar}>='" & Text1.Text & "' And {tblbayarspp.tglbayar}<= '" & Text2.Text & "'"
        .Formulas(0) = "Awal=' Dari Tanggal: " & Text1.Text & "'"
        .Formulas(1) = "Akhir=' Sampai Tanggal: " & Text2.Text & "'"
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .Action = 1
        
    End With
Me.MousePointer = 1
End Sub
Private Sub CetakNunggak()
Me.MousePointer = 11
With crtCetak
        .Reset
        .ReportFileName = App.Path & "\LapNunggak.rpt"
        '.Password = Chr(10) & "amir"
        .DataFiles(0) = App.Path & "\madrasah.mdb"
        .SelectionFormula = "{tblnunggak.sppbulan}='" & Combo1.Text & "' And {tblnunggak.tahun}= '" & Text3.Text & "'"
        .Formulas(0) = "Bulan=' Bulan: " & Combo1.Text & "'"
        .Formulas(1) = "Tahun=' Tahun: " & Text3.Text & "'"
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .Action = 1
        
    End With
Me.MousePointer = 1
End Sub

Private Sub Command3_Click()
On Error Resume Next
If opt1 = True Then
CetakSPP
End If
If opt2 = True Then
CetakNunggak
End If
End Sub


Private Sub command4_Click()
frmcarisiswa.Show
Command5.Enabled = True
End Sub

Private Sub command5_Click()
Me.MousePointer = 11
With crtCetak
        .Reset
        .ReportFileName = App.Path & "\LapPemSis.rpt"
        '.Password = Chr(10) & "amir"
        .DataFiles(0) = App.Path & "\madrasah.mdb"
        .SelectionFormula = "{tblbayarspp.nama}='" & TxtNama.Text & "' And {tblbayarspp.tglbayar}>='" & DTPicker1.Value & "' And {tblbayarspp.tglbayar}<= '" & DTPicker2.Value & "'"
        .Formulas(0) = "Awal=' Dari Tanggal: " & DTPicker1.Value & "'"
        .Formulas(1) = "Akhir=' Sampai Tanggal: " & DTPicker2.Value & "'"
        .Formulas(2) = "NamaSiswa= 'Nama Siswa : " & TxtNama.Text & " '"
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .Action = 1
        
    End With
Me.MousePointer = 1

End Sub

Private Sub Form_Load()
ada = True
Data1.DatabaseName = App.Path & "\madrasah.mdb"
Data1.RecordsetType = 1
Data2.DatabaseName = App.Path & "\madrasah.mdb"
Data2.RecordsetType = 1
Data3.DatabaseName = App.Path & "\madrasah.mdb"
Data3.RecordsetType = 1
Data3.RecordSource = "tblnunggak"
Data4.DatabaseName = App.Path & "\madrasah.mdb"
Data4.RecordsetType = 0
Data4.RecordSource = "tblsiswa"
Command3.Enabled = False
Text1.Text = Date
Text2.Text = Date
Text3.Text = Year(Date)
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
Command5.Enabled = False
DBGrid1.Visible = False
DBGrid2.Visible = False
DTPicker1 = Format(Date, "dd MM yyyy")
DTPicker2 = Format(Date, "dd MM yyyy")
End Sub

Private Sub opt1_Click()
Label7.Caption = "Daftar Siswa Yang Membayar"
End Sub

Private Sub opt2_Click()
Label7.Caption = "Daftar Siswa Yang Menunggak"
End Sub

Private Sub Text1_LostFocus()
If Text1.Text = Empty Then
   Text1.Text = Date
End If

End Sub
Private Sub Text2_LostFocus()
If Text2.Text = Empty Then
   Text2.Text = Date
End If

End Sub
Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then  'perintah pada saat penekanan tombol enter
   Text2.SetFocus
End If
End Sub
Private Sub text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then  'perintah pada saat penekanan tombol enter
   Command1.SetFocus
End If
End Sub

Private Sub TxtNis_Change()
With Data4.Recordset
      .Index = "nisx"
      .Seek "=", TxtNis.Text
      If Not .NoMatch Then
         tampilkan
        Else
        TxtNama.Text = ""
         End If
   End With
End Sub

Private Sub tampilkan()
With Data4.Recordset
  TxtNama.Text = !nama
End With
End Sub
