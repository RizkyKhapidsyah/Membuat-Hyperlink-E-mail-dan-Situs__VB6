VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuat Hyperlink E-mail dan Situs"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblSitus 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label lblEmail 
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jika Anda ingin agar ketika mouse berada di atas kedua 'label tersebut berubah menjadi kursor bergambar 'tangan, set property MouseIcon saat design-time dengan 'file icon bergambar tangan yang Anda miliki, dan
'property MousePointer menjadi "99 - Custom".

'Deklarasikan fungsi API untuk mengeksekusi suatu 'Hyperlink
Private Declare Function ShellExecute _
Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd _
As Long, ByVal lpOperation As String, ByVal lpFile _
As String, ByVal lpParameters As String, ByVal _
lpDirectory As String, ByVal nShowCmd As Long) As Long

Const SW_SHOWNORMAL = 1  'Konstanta untuk menampilkan 'jendela normal

Private Sub Form_Load()  'Tampilkan nama untuk setiap 'label
  lblEmail.Caption = "rizkykhapidsyah@gmail.com"  'Email
  lblEmail.Font.Underline = True
  lblEmail.ForeColor = vbBlue  'Warna link mula-mula biru
  lblSitus.Caption = "https://programmerfiles.blogspot.com/" 'Situs
  lblSitus.Font.Underline = True
  lblSitus.ForeColor = vbBlue  'Warna link mula-mula
                               'biru
End Sub

Private Sub lblEmail_Click()
  Dim email As Long
  'Tampilkan program default untuk mengirim e-mail ke
  'alamat lblEmail
  email = ShellExecute(0, vbNullString, _
         "mailto:" & lblEmail, "", "", vbNormalFocus)
  lblEmail.ForeColor = &H8000& 'Setelah diklik, berubah
                               'warna
End Sub

Private Sub lblSitus_Click()
  Dim situs As Long
  'Tampilkan program default untuk membuka situs ke
  'alamat lblSitus
  situs = ShellExecute(0, vbNullString, _
          lblSitus, "", "", vbNormalFocus)
  lblSitus.ForeColor = &H8000& 'Setelah diklik, berubah
                               'warna
End Sub


