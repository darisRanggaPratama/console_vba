VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuatWorksheet 
   Caption         =   "Buat Worksheet"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3930
   OleObjectBlob   =   "BuatWorksheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BuatWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Kode saat opsi Default dipilih
Private Sub optDefault_Click()

'TextBox Nama menjadi tidak aktif
txtNama.Enabled = False

End Sub

'Kode saat opsi Ubah dipilih
Private Sub optUbah_Click()

'TextBox Nama menjadi aktif
txtNama.Enabled = True

End Sub

'Kode apabila tombol OK ditekan
Private Sub cmdOK_Click()

'Lanjutkan Macro jika terjadi error
On Error Resume Next

'Membuat worksheet baru
Set NewSheet = Application.Sheets.Add

'Jika opsi Default dalam keadaan terpilih
If optDefault.Value = True Then
    'Keluar dari Sub Prosedur
    Exit Sub
'Jika opsi Ubah dalam keadaan terpilih
ElseIf optUbah.Value = True Then
    'Jika nama worksheet yang ditulis belum ada
    If Application.Sheets(txtNama.Text) Is Nothing Then
        'Memberi nama worksheet
        NewSheet.Name = txtNama.Text
        'Menonaktifkan kotak dialog
        Unload Me
    Else
        'Menonaktifkan kotak dialog
        Unload Me
        'Kotak pesan jika nama worksheet sudah ada
        MsgBox "Nama worksheet sudah ada", _
            vbOKOnly + vbInformation
    End If
End If

End Sub

'Kode saat tombol Batal ditekan
Private Sub cmdBatal_Click()

'Menonaktifkan kotak dialog
Unload Me

End Sub
