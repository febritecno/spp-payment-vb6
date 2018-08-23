VERSION 5.00
Begin VB.Form Flogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2805
   ClientLeft      =   15
   ClientTop       =   60
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton TombolBtl 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton TombolOK 
      Caption         =   "LOGIN"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
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
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN PROGRAM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama User"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Flogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TombolBtl_Click()
End
End Sub

Private Sub TombolOK_Click()
If Text1.Text = "admin" And Text2.Text = "admin" Then
   Fmenuutama.menuaktif
   Fmenuutama.Show
   Unload Me
Else
   MsgBox "User dan password anda salah", 0 + 16, "Konfirmasi"
   Exit Sub
End If
End Sub
