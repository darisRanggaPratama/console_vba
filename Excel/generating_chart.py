import pandas as pd
import matplotlib.pyplot as plt
import sys
from openpyxl import load_workbook

def generate_charts(excel_file):
    # Load data from Excel
    data = pd.read_excel(excel_file, sheet_name='DummyData')
    
    # Create figure with subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 12))
    
    # Plot 1: Line chart for products
    ax1.plot(data['Tahun'], data['Jumlah Produk'], 'b-o', linewidth=2)
    ax1.set_title('Jumlah Produk per Tahun')
    ax1.set_xlabel('Tahun')
    ax1.set_ylabel('Jumlah Produk')
    ax1.grid(True)
    
    # Plot 2: Bar chart for materials
    ax2.bar(data['Tahun'], data['Jumlah Material'], color='green')
    ax2.set_title('Jumlah Material per Tahun')
    ax2.set_xlabel('Tahun')
    ax2.set_ylabel('Jumlah Material')
    ax2.grid(True, axis='y')
    
    # Adjust layout and save
    plt.tight_layout()
    plt.savefig(excel_file.replace('.xlsm', '.png').replace('.xlsx', '.png'))
    plt.savefig(excel_file.replace('.xlsm', '.pdf').replace('.xlsx', '.pdf'))
    plt.show()

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python generate_chart.py <excel_file>')
        sys.exit(1)
    
    excel_file = sys.argv[1]
    generate_charts(excel_file)
