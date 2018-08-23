VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FBayarSPP 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entry Bayar SPP"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   ControlBox      =   0   'False
   Icon            =   "FBayarSPP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Tombol1 
      Caption         =   "CREATE"
      Height          =   735
      Left            =   7680
      TabIndex        =   36
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "SIMPAN"
      Height          =   735
      Left            =   7680
      TabIndex        =   35
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Tombol2 
      Caption         =   "CEK NOTA"
      Height          =   735
      Left            =   7680
      TabIndex        =   34
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "HAPUS"
      Height          =   735
      Left            =   7680
      TabIndex        =   33
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "BATAL"
      Height          =   735
      Left            =   7680
      TabIndex        =   32
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "KELUAR"
      Height          =   735
      Left            =   7680
      TabIndex        =   31
      Top             =   4440
      Width           =   1815
   End
   Begin Crystal.CrystalReport crtNota 
      Left            =   6600
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   6600
      Width           =   615
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.TextBox Text11 
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
      Left            =   4800
      TabIndex        =   29
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Input Data Siswa Yang Menunggak"
      Height          =   375
      Left            =   1920
      TabIndex        =   26
      Top             =   8160
      Width           =   5655
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   7335
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   23
      ToolTipText     =   "mengikuti jam sistem"
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "mengikuti jam sistem"
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
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
      Height          =   405
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "mengikuti tanggal sistem"
      Top             =   240
      Width           =   2055
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
      Left            =   960
      TabIndex        =   2
      Top             =   2595
      Width           =   2415
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "mengikuti jam sistem"
      Top             =   3360
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
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
      Height          =   405
      Left            =   5760
      TabIndex        =   14
      ToolTipText     =   "mengikuti tanggal sistem"
      Top             =   210
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FBayarSPP.frx":0442
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "FBayarSPP.frx":0456
      TabIndex        =   7
      Top             =   5520
      Width           =   9375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   7335
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
         Left            =   1800
         MaxLength       =   75
         TabIndex        =   17
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
         Left            =   1800
         MaxLength       =   75
         TabIndex        =   16
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   300
         Left            =   3720
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
         Left            =   1800
         MaxLength       =   75
         TabIndex        =   5
         Top             =   600
         Width           =   5175
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.Data Data1 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5880
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
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Siswa"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NIS"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Index           =   2
      Left            =   4200
      TabIndex        =   28
      Top             =   2640
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terbilang"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   3960
      Width           =   660
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Uang Kembali"
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Yang dibayar"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   1470
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Kwitansi"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "SPP Bulan"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
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
      TabIndex        =   19
      Top             =   7800
      Width           =   5175
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
      TabIndex        =   18
      Top             =   5280
      Width           =   9375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Biaya SPP"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Bayar"
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FBayarSPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ada As Boolean
Dim NoFaktur As Integer

'untuk menonaktifkan tombol command
Private Sub tomboltidakaktif()
cmdHapus.Enabled = False
End Sub
'Untuk mengaktikan tombol command
Private Sub tombolaktif()
cmdHapus.Enabled = True
End Sub

Private Sub Command1_Click()
frmcarisiswa.Show 1
End Sub

Private Sub Command2_Click()
Form1.Show
Unload Me
End Sub





Private Sub DBGrid1_Click()
cmdHapus.Enabled = True
End Sub

Private Sub Text1_Change()
With Data1.Recordset
      .Index = "nisx"
      .Seek "=", Text1.Text
      If Not .NoMatch Then
         tampilkan
         textboxtdkaktif
         tombolaktif
         cmdSimpan.Enabled = True
      Else
         kosong
         cmdSimpan.Enabled = False
      End If
   End With
End Sub

'perintah penekanan tombol enter
Private Sub text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text5.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
If Text3.Text = Empty Then
   Text3.Text = Format(Date, "yyyy-mm-dd")
   End If
End Sub

Private Sub Text4_Change()
Text10.Text = BuatTerbilang(Val(Text4.Text))
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)

If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or _
   KeyAscii = vbKeyBack) Then
   KeyAscii = 0
End If
If KeyAscii = 13 Then  'perintah pada saat penekanan tombol enter
   Text8.SetFocus
End If
End Sub

Private Sub Text8_Change()
On Error Resume Next
Text9.Text = Val(Text8.Text) - Val(Text4.Text)
End Sub

Private Sub text8_KeyPress(KeyAscii As Integer)

If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or _
   KeyAscii = vbKeyBack) Then
   KeyAscii = 0
End If
If KeyAscii = 13 Then  'perintah pada saat penekanan tombol enter
   cmdSimpan.SetFocus
End If
End Sub
Private Sub text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text6.SetFocus
End If
End Sub
Private Sub text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdSimpan.SetFocus
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
Private Sub cmdBatal_Click()
tomboltidakaktif
textboxtdkaktif
kosong
Text1.Enabled = True
Text1.SetFocus
End Sub
Private Sub cmdHapus_Click()
cmdHapus.Enabled = False
Data3.RecordSource = "Select * From tblbayarspp Where nis = '" & DBGrid1.Text & "'"
Data3.Refresh
With Data3.Recordset
     If Not .EOF Then
         Pesan = MsgBox("Yakin data ini akan dihapus ...?", vbYesNo, "Konfirmasi")
         If Pesan = vbYes Then
           .Delete
           Data3.Refresh
           'Data2.Refresh
           cmdBatal_Click
           cmdSimpan.Enabled = True
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
'CreateNoFaktur
If (Text1.Text = Empty Or Combo1.Text = Empty Or Text4.Text = Empty Or Text7.Text = Empty) Then
      MsgBox "Data belum lengkap !", , "Konfirmasi"
      Exit Sub
Else ' jika sudah lengkap
   
   Data3.RecordSource = "Select * From tblbayarspp Where nis = '" & Text1.Text & "' and sppbulan = '" & Combo1.Text & "'"
   Data3.Refresh
   With Data3.Recordset
      If .EOF Then
           .AddNew
           !nokwitansi = Text7.Text
           !tglbayar = Text3.Text
           !nis = Text1.Text
           !nama = Text2.Text
           !sppbulan = Combo1.Text
           !biayaspp = Text4.Text
           !jumlahbayar = Text8.Text
           !ukem = Text9.Text
           !terbilang = Text10.Text
           !tahun = Left(Text3, 4)
           !Status = Text11.Text
           .Update
           Text8.Text = ""
           Text9.Text = ""
      Else
          MsgBox "Siswa tersebut sudah membayar untuk spp bulan tersebut", 0 + 16, "Info"
          kosong
          Exit Sub
      End If
    End With
End If
Data3.Refresh
'Data2.Refresh
kosong
cmdBatal_Click
End Sub
'untuk mengosongkan kembali textbox
Private Sub kosong()
Text5.Text = "": Text1.Text = "": Text2.Text = "": Text6.Text = "": Text4.Text = "": Combo1.Text = "": Text11.Text = "": Text10.Text = "": Text9.Text = "": Text8.Text = ""
End Sub

'perintah pada saat form dijalankan
Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\madrasah.mdb"
Data3.DatabaseName = App.Path & "\madrasah.mdb"
Data1.RecordsetType = 0
Data3.RecordsetType = 1
Data1.RecordSource = "tblsiswa"
Data3.RecordSource = "tblbayarspp"
Text3.Text = Format(Date, "yyyy-mm-dd")
'CreateNoFaktur
'Data3.RecordSource = "Select tblbayarspp.nokwitansi,tblbayarspp.nis,tblsiswa.nama,tblsiswa.kelas,tblbayarspp.sppbulan,tblbayarspp.biayaspp,tblbayarspp.terbilang From tblbayarspp,tblsiswa Where tblbayarspp.nis=tblsiswa.nis and cdate(tblbayarspp.tglbayar) = '" & Format(Text3.Text, "dd/mm/yyyy") & "' order by tblbayarspp.nis"
Data3.Refresh
tomboltidakaktif
textboxtdkaktif
ada = True
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
Label7.Caption = "Daftar Siswa Yang Membayar Tanggal : " + Text3.Text
cmdSimpan.Enabled = False
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
         cmdSimpan.Enabled = True
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
  Text5.Text = !jnskelamin
  Text6.Text = !kelas
  Text4.Text = !spp
  Text11.Text = !Status
End With
End Sub
'Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = vbKeyReturn Then
'     SendKeys vbTab
'  End If
  
'End Sub
Private Sub Text8_LostFocus()
If Text8.Text = Empty Then
 Exit Sub
 End If
If Val(Text8.Text) < Val(Text4.Text) Then
   MsgBox "Jumlah bayar kurang, silahkan input kembali", 0 + 16, "Konfirmasi"
   Text8.Text = ""
   Text8.SetFocus
   Exit Sub
End If
End Sub
Private Sub CreateNoFaktur()
Data3.Refresh
With Data3.Recordset
If .EOF = False Then
.MoveFirst
Do While Not .EOF
Text7.Text = .Fields("nokwitansi")
.MoveNext
Loop
NoFaktur = NoFaktur + 1
Text7.Text = Format(Date, "yymmdd") & Format(NoFaktur, "000")
Else
Text7.Text = Format(Date, "yymmdd") & Format(NoFaktur, "000")
End If
End With
End Sub

Private Sub Tombol1_Click()
CreateNoFaktur
End Sub

Sub CetakNota()
    crtNota.ReportFileName = App.Path & "\Kwitansi.rpt"
    crtNota.WindowState = crptMaximized
    crtNota.RetrieveDataFiles
    crtNota.Action = 1
End Sub

Private Sub Tombol2_Click()
CetakNota
End Sub
