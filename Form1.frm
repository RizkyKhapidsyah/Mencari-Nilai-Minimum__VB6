VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencari Nilai Minimum"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim arrData(10) As Integer
Dim Min As Integer, i As Integer
  'Isi elemen array arrData
  arrData(0) = 12
  arrData(1) = 500
  arrData(2) = 92
  arrData(3) = 262
  arrData(4) = 112
  arrData(5) = 152
  arrData(6) = 887
  arrData(7) = 10
  arrData(8) = 120
  arrData(9) = 12

  'Inisialisasi variabel Min
  Min = 32767
  
  'Mengapa 32767...? Tentu Anda bertanya demikian?
  'Jawabnya: Karena tipe data yang kita gunakan untuk
  '          variabel Min adalah Integer, dan memiliki
  '          jangkauan nilai dari - 32768 sampai dengan
  '          32767. Nilai ini bisa Anda perbesar jika
  '          Anda menggunakan tipe data Long
  '          (misalnya).
  'Lalu Anda mungkin bertanya lagi: Mengapa nilai
  'terbesar diambil? Jawabnya: Karena kita akan
  'membandingkan dengan nilai lainnya yang belum
  'diketahui, jadi kita inisialisasikan nilai variabel
  'pembanding dengan nilai terbesar dari range
  'tipe data yang kita gunakan. Bandingkan dengan tips
  '"Cari Nilai Maksimum".

  'Bersihkan form
  Form1.Cls
  'Periksa semua isi array
  For i = 0 To 9
    'Cetak data-nya ke layar
    Print arrData(i)
    'Jika array indeks ke-i lebih kecil dari Min
    If arrData(i) < Min Then
       'Tampung nilai Min
       Min = arrData(i)
    Else 'Jika tidak...
       'Nilai Min masih tetap yang sebelumnya
       Min = Min
    End If 'Akhir pemeriksaan isi array
  Next i
  'Tampilkan nilai minimum setelah selesai iterasi
  MsgBox "Nilai minimum = " & Min, _
         vbInformation, "Minimum"
End Sub



