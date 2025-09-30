Attribute VB_Name = "ArrColl"
Option Explicit

Public Sub RunDemo()
    texts ("-----")
    texts ("Run Demo: Perbandingan Array vs Collection")
    texts ("Tanggal/Time (Timer): " & Timer)
    texts ("-----")

    DemoArray


End Sub

Public Sub texts(ByVal message As String)
    Debug.Print message
End Sub

' Demonstrasi Array
Public Sub DemoArray()
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    Dim posToRemove As Long

    texts ("- Demo: Array -")

    ' - Fixed-size Array -
    Dim fixedArr(0 To 4) As String
    texts ("Fixed-size array dideklarasikan: fixedArr(0 To 4) -> total 5 elemen")

    ' Isi data via indeks (akses cepat)
    fixedArr(0) = "Alpha"
    fixedArr(1) = "Bravo"
    fixedArr(2) = "Charlie"
    fixedArr(3) = "Delta"
    fixedArr(4) = "Echo"

    texts ("Isi fixedArr (index 0..4):")
    For i = LBound(fixedArr) To UBound(fixedArr)
        texts ("  fixedArr(" & i & ") = " & fixedArr(i))
    Next i

    ' - Contoh Error: Akses di luar batas indeks -
    Err.Clear
    fixedArr(5) = "OverFlow" ' Mencoba menulis index di luar batas -> runtime error 9
    If Err.Number <> 0 Then
        texts(" [Error] saat menulis fixedArr(5): Err.Number = " & Err.Number & ", Desc = " & Err.Description)
        Err.Clear
    Else
        texts(" fixedArr(5) berhasil (seharusnya tidak).")
    End If

    ' - Dynamic Array + ReDim Preserve -
    Dim dynArr() As Long
    Redim dynArr(1 To 3) ' dynamic with initial size
    dynArr(1) = 10
    dynArr(2) = 20
    dynArr(3) = 30
    texts("Dynamic array awal: ReDim dynArr(1 To 3)")
    For i = LBound(dynArr) To UBound(dynArr)
        texts("  dynArr(" & i & ") = " & dynArr(i))
    Next i

    ' Resize: memperbesar dengan ReDim Preserve (menjaga isi)
    ReDim Preserve dynArr(1 To 5)
    dynArr(4) = 40
    dynArr(5) = 50
    texts("After ReDim Preseve dynArr(1 To 5):")
    For i = LBound(dynArr) To UBound(dynArr)
        texts("  dynArr(" & i & ") = " & dynArr(i))
    Next i

    ' - Menghapus elemen di Array (workaround) -
    texts("Contoh: menghapus elemen posisi 2 (bukan built-in; gunakan shift + ReDim Preserve)")
    posToRemove = 2
    texts(" Sebelum penghapusan (count = " & (UBound(dynArr) - LBound(dynArr) + 1) & "):")
    For i = LBound(dynArr) To UBound(dynArr)
        texts("  dynArr(" & i & ") = " & dynArr(i))
    Next i

    ' Shift left dari posisi penghapusan
    For i = posToRemove to UBound(dynArr) - 1
        dynArr(i) = dynArr(i + 1)
    Next i

    ' Shrink array dengan ReDim Preserve
    ReDim Preserve dynArr(LBound(dynArr) To UBound(dynArr) - 1)

    Texts (" Setelah penghapusan posisi " & posToRemove & " (count sekarang = " & (UBound(dynArr) - LBound(dynArr) + 1) & "):")
    For i = LBound(dynArr) To UBound(dynArr)
        texts("  dynArr(" & i & ") = " & dynArr(i))
    Next i

    '- Iterasi: For(index) vs For Each -
    texts("Iterasi array: For i = LBound To UBound (umum & cepat):")
    For i = LBound(dynArr) To UBound(dynArr)
        texts("  index " & i & " -> " & dynArr(i))
    Next i

    texts("Iterasi array: For Each (bisa digunakan pada array):")
    Dim V As Variant
    For Each V In dynArr
        texts("  value -> " & V)
    Next V

    texts("Catatan: Array - ukuran bisa tetap (fixed) atau diubah dengan ReDim (dynamic), akses via index cepat.")
End Sub

' Demonstrasi Collection
Public sub DemoCollection()
    on error resume next
    texts(" - DEMO: Collection -")

    dim coll as new collection
    dim i as long
    dim itm as variant

    ' Tambah elemen (dinamis)
    coll.add "Satu", "kSatu" ' dengan key
    coll.add "Dua" ' tanpa key (akses via index)
    coll.add "Tiga", "kTiga"
    coll.add "Empat" ' tanpa key
    texts("Collection ditambahkan 4 elemen (beberapa dengan Key): count = " & coll.count)

    ' Akses by index dan by key
    texts("Akses by index coll(2) = " & coll(2))
    texts("Akses by key coll(""kSatu"") = " & coll("kSatu"))

    ' Coba akses key yang tidak ada -> akan munculkan error (ditangani)
    err.clear
    dim temp as variant
    temp = coll("tidakAdaKey") ' runtime error 5 (invalid procedure call or argument)
    if err.number <> 0 then
        texts(" [Error] Akses key yang tidak ada: Err.Number = " & err.number & ", Desc = " & err.Description)
        err.clear
    else
        texts(" coll(""tidakAdaKey"") = " & temp)
    end if

    ' Hapus elemen: Remove by index dan Remove by key
    texts("Hapus elemen index 2 (coll.Remove 2)")
    err.clear
    coll.remove 2 ' runtime error 5 (invalid procedure call or argument)
    if err.number <> 0 then
        texts(" [Error] Saat remove by index: " & err.number & " - " & err.Description)
        err.clear
    else
        texts(" Berhasil remove index 2. count sekarang = " & coll.count)
    end if

    texts("Hapus elemen by key 'kSatu' (coll.Remove ""kSatu"")")
    err.clear
    coll.remove "kSatu"
    if err.number <> 0 then
        texts(" [Error] saat Remove by key: " & err.number & " - " & err.Description)
        err.clear
    else
        texts(" Berhasil remove key 'kSatu'. count sekarang = " & coll.count)
    end if

    ' Iterasi: For Each (disukai untuk Collection) vs For i = 1 To coll.Count
    texts("Iterasi Collection: For Each (preferred):")
    for each itm in coll
        texts(" For Each -> " & itm)
    next itm

    texts("Iterasi Collection: For i = 1 To coll.Count (akses via coll(i)):")
    for i = 1 to coll.count
        texts(" coll(" & i & ") = " & coll(i))
    next i

    ' Menunjukkan penggunaan Key lebih intuitif:
    

End Sub


