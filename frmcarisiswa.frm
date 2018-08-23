VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmcarisiswa 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari siswa"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2340
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
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmcarisiswa.frx":0000
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "frmcarisiswa.frx":0014
      TabIndex        =   2
      Top             =   720
      Width           =   7455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NIS Aktif"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Klik 2x pada kolom nis untuk memilih data siswa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ketik nama siswa"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmcarisiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_Click()
Text2.Text = DBGrid1.Text
End Sub

Private Sub DBGrid1_DblClick()
On Error Resume Next
If Fsiswa.ada = True Then
   Fsiswa.Text1.Text = frmcarisiswa.Text2.Text
   Unload Me
End If
If Form1.ada = True Then
   Form1.Text1.Text = frmcarisiswa.Text2.Text
   Unload Me
End If
If FBayarSPP.ada = True Then
   FBayarSPP.Text1.Text = frmcarisiswa.Text2.Text
   Unload Me
End If
If FLaporan.ada = True Then
   FLaporan.TxtNis.Text = frmcarisiswa.Text2.Text
   Unload Me
End If
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\madrasah.mdb"
Data1.RecordsetType = 1
Data1.RecordSource = "Select nis,nama From tblsiswa order by nis"
End Sub

Private Sub Text1_Change()
'On Error Resume Next
Data1.RecordSource = "Select nis,nama From tblsiswa Where nama like '*" & Text1.Text & "*' order by nis"
Data1.Refresh
End Sub
