import xlwings as xw
import matplotlib.pyplot as plt

def create_graph():
    # Menghubungkan ke workbook Excel yang sedang aktif
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    
    # Mengambil data dari Excel (baris 2 sampai 6, sesuai data dummy)
    years = sheet.range('A2:A6').value
    products = sheet.range('B2:B6').value
    materials = sheet.range('C2:C6').value
    
    # Membuat grafik
    plt.figure(figsize=(10, 6))
    plt.plot(years, products, label='Jumlah Produk', marker='o')
    plt.plot(years, materials, label='Jumlah Material', marker='s')
    plt.xlabel('Tahun')
    plt.ylabel('Jumlah')
    plt.title('Grafik Jumlah Produk dan Material Selama 5 Tahun')
    plt.legend()
    plt.grid(True)
    
    # Menampilkan grafik
    plt.show()

if __name__ == "__main__":
    xw.serve()