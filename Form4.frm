VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H80000005&
   Caption         =   "SURAT KELUAR"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   2055
      Left            =   720
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   1251.765
      ScaleMode       =   0  'User
      ScaleTop        =   500
      ScaleWidth      =   1419.355
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CETAK LAPORAN"
      Height          =   855
      Left            =   17760
      TabIndex        =   26
      Top             =   720
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   7560
      Top             =   7680
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\ASUS\Documents\latihan\Latihan2.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\ASUS\Documents\latihan\Latihan2.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from login2"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Kembali Ke Menu"
      Height          =   735
      Left            =   17160
      TabIndex        =   25
      Top             =   9480
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   6255
      Begin VB.TextBox Text1 
         DataField       =   "Tanggal"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   1680
         TabIndex        =   14
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         DataField       =   "NoSurat Keluar"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         DataField       =   "Tgl Surat"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox Text4 
         DataField       =   "Terima Dari"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox Text5 
         DataField       =   "Perihal"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   3120
         Width           =   3855
      End
      Begin VB.TextBox Text6 
         DataField       =   "Lampiran"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   3840
         Width           =   3855
      End
      Begin VB.TextBox Text7 
         DataField       =   "Ket"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   4440
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "NoSurat Keluar"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Tgl Surat"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Terima Dari"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Perihal"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Lampiran"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   15
         Left            =   -2160
         TabIndex        =   16
         Top             =   -1080
         Width           =   15
      End
      Begin VB.Label Label9 
         Caption         =   "Ket"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   4560
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   9600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   9600
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":2614
      Height          =   5655
      Left            =   7560
      TabIndex        =   1
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9975
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "SURAT   KELUAR"
      BeginProperty Font 
         Name            =   "Poplar Std"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   24
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      Caption         =   "LIST     SURAT   KELUAR"
      BeginProperty Font 
         Name            =   "Poplar Std"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11400
      TabIndex        =   23
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Adodc1.Recordset.AddNew
End Sub

Private Sub Keluar_akun_Click()
Form1.Show
Unload Me
End Sub

Private Sub Command10_Click()
Form2.Show
Unload Me
End Sub


Private Sub Command2_Click()
Text1.SetFocus
Command2.Caption = "&Edit"
End Sub


Private Sub Command3_Click()
If MsgBox("Yakin Ingin Menghapus Data?", vbQuestion + vbOKCancel, "konfirmasi") = vbOK Then
Adodc1.Recordset.Delete
Me.DataGrid1.Refresh
End If
End Sub


Private Sub Command4_Click()
Dim a
a = MsgBox("Apakah Data Ingin Disimpan...?", vbQuestion + vbYesNo)
If a = vbYes Then
MsgBox "Data Tersimpan...!", vbInformation, "Pesan"
Else
Exit Sub
End If
End Sub

Private Sub Command5_Click()
Dim a As String
a = MsgBox("Apakah Anda Ingin Keluar", vbYesNo + vbQuestion, "Perhatian")
If a = vbYes Then Unload Me
End Sub

Private Sub Command6_Click()
Form6.Show
'Menampilkan Form6 ketika Command Button diklik.
Unload Me
'Menutup Form4 ketika Form6 terbuka
End Sub
