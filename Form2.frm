VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000005&
   Caption         =   "MENU APLIKASI"
   ClientHeight    =   8520
   ClientLeft      =   525
   ClientTop       =   1725
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   ScaleHeight     =   8520
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000005&
      Caption         =   "KELUAR AKUN"
      Height          =   1935
      Left            =   16440
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Caption         =   "Laporan"
      Height          =   3015
      Left            =   15840
      TabIndex        =   6
      Top             =   360
      Width           =   3855
      Begin VB.CommandButton Command6 
         BackColor       =   &H80000005&
         Caption         =   "LAPORAN SURAT KELUAR"
         Height          =   2055
         Left            =   2040
         Picture         =   "Form2.frx":0FD6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000005&
         Caption         =   "LAPORAN SURAT MASUK"
         Height          =   2055
         Left            =   360
         Picture         =   "Form2.frx":1D5E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "Tentang"
      Height          =   3015
      Left            =   7800
      TabIndex        =   3
      Top             =   240
      Width           =   3735
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000005&
         Caption         =   "KONTAK"
         Height          =   2055
         Left            =   1920
         Picture         =   "Form2.frx":2AE7
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000005&
         Caption         =   "TENTANG APP"
         Height          =   2055
         Left            =   360
         Picture         =   "Form2.frx":38DC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000005&
      Caption         =   "SURAT KELUAR"
      Height          =   2055
      Left            =   2040
      Picture         =   "Form2.frx":4789
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "SURAT MASUK"
      Height          =   2055
      Left            =   600
      Picture         =   "Form2.frx":56B1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Menu"
      Height          =   3015
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000005&
      Caption         =   "Akun"
      Height          =   3495
      Left            =   16080
      TabIndex        =   10
      Top             =   4800
      Width           =   3855
      Begin VB.CommandButton Command8 
         BackColor       =   &H80000005&
         Caption         =   "KELUAR APP"
         Height          =   1935
         Left            =   2040
         Picture         =   "Form2.frx":61F9
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Surat_Masuk 
         Caption         =   "Surat Masuk"
      End
      Begin VB.Menu Surat_Keluar 
         Caption         =   "Surat Keluar"
      End
   End
   Begin VB.Menu Tentang 
      Caption         =   "Tentang"
      Begin VB.Menu ttgapps 
         Caption         =   "Tentang Apps"
      End
      Begin VB.Menu Kontak 
         Caption         =   "Kontak"
      End
   End
   Begin VB.Menu laporan 
      Caption         =   "Laporan"
      Begin VB.Menu laporansuratmasuk 
         Caption         =   "Laporan Surat Masuk"
      End
      Begin VB.Menu laporansuratkeluar 
         Caption         =   "Laporan Surat Keluar"
      End
   End
   Begin VB.Menu Akun 
      Caption         =   "Akun"
      Begin VB.Menu Keluar_akun 
         Caption         =   "Keluar akun"
      End
      Begin VB.Menu Keluar_aplikasi 
         Caption         =   "Keluar aplikasi"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form3.Show
'Menampilkan Form3 ketika Command Button diklik.
Unload Me
'Menutup Form2 ketika Form3 terbuka
End Sub

Private Sub Command2_Click()
Form4.Show
'Menampilkan Form4 ketika Command Button diklik.
Unload Me
'Menutup Form3 ketika Form4 terbuka
End Sub

Private Sub Command3_Click()
X = MsgBox("dibuat oleh developer", 0, "Tentang Developer")
End Sub

Private Sub Command4_Click()
X = MsgBox("vb06amikbgr@ymail.com", 0, "Kontak")
End Sub

Private Sub Command5_Click()
Form5.Show
'Menampilkan Form4 ketika Command Button diklik.
Unload Me
'Menutup Form3 ketika Form5 terbuka
End Sub

Private Sub Command6_Click()
Form6.Show
'Menampilkan Form6 ketika Command Button diklik.
Unload Me
'Menutup Form4 ketika Form6 terbuka
End Sub

Private Sub Command7_Click()
Form1.Show
Unload Me
End Sub

Private Sub Command8_Click()
Dim a As String
a = MsgBox("Apakah Anda Ingin Keluar", vbYesNo + vbQuestion, "Perhatian")
If a = vbYes Then Unload Me
End Sub

Private Sub Keluar_akun_Click()
Form1.Show
Unload Me
End Sub

Private Sub Keluar_aplikasi_Click()
Dim a As String
a = MsgBox("Apakah Anda Ingin Keluar", vbYesNo + vbQuestion, "Perhatian")
If a = vbYes Then Unload Me
End Sub

Private Sub Kontak_Click()
X = MsgBox("vb06amikbgr@ymail.com", 0, "Kontak")
End Sub


Private Sub Label1_Click()

End Sub

Private Sub laporansuratmasuk_Click()
Form5.Show
'Menampilkan Form4 ketika Command Button diklik.
Unload Me
'Menutup Form3 ketika Form5 terbuka
End Sub

Private Sub Surat_Keluar_Click()
Form4.Show
'Menampilkan Form4 ketika Command Button diklik.
Unload Me
'Menutup Form3 ketika Form4 terbuka
End Sub

Private Sub Surat_Masuk_Click()
Form3.Show
'Menampilkan Form3 ketika Command Button diklik.
Unload Me
'Menutup Form2 ketika Form3 terbuka
End Sub

Private Sub Timer1_Timer()
Label1.ForeColor = RGB(Rnd * 250, Rnd * 250, Rnd * 250)
    If (Label1.Left + Label1.Width) <= 0 Then
        Label1.Left = Me.Width
    End If
    Label1.Left = Label1.Left - 100
End Sub

Private Sub ttgapps_Click()
X = MsgBox("dibuat oleh developer", 0, "Tentang Developer")
End Sub
