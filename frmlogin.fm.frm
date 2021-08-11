VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   Caption         =   "~LOGIN~"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000005&
      Caption         =   "&KELUAR"
      Height          =   975
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmlogin.fm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "&LOGIN"
      Height          =   975
      Left            =   1200
      Picture         =   "frmlogin.fm.frx":0DB7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtpassword 
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtusername 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "FORM LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call konekdb
If txtusername = "admin" And txtpassword = "1234" Then
X = MsgBox("Hello Admin", 0, "Congratulations !")
Form2.Show 'Perintah Menampilkan Form 2
Form1.Visible = False 'Menyembunyikan Form 1
Unload Me 'Menutup Form 1
Else
MsgBox "User Name atau Password yang Anda Masukkan salah" _
& vbNewLine & "Silahkan Coba lagi !!", vbCritical, "Warning!!"
txtusername = ""
txtpassword = ""
txtusername.SetFocus
End If
End Sub
Private Sub Command2_Click()
Dim a As String
a = MsgBox("Apakah Anda Ingin Keluar", vbYesNo + vbQuestion, "Perhatian")
If a = vbYes Then Unload Me
End Sub

