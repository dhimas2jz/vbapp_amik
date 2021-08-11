VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H8000000E&
   Caption         =   "CETAK LAPORAN SURAT KELUAR"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form6"
   ScaleHeight     =   6135
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK LAPORAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KEMBALI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   0
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "ANDA YAKIN MAU MENCETAK LAPORAN ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    DataReport2.Show 1
End Sub

Private Sub Command2_Click()
Form3.Show
Unload Me
End Sub
