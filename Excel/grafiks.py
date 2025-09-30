import matplotlib.pyplot as plt
import numpy as np
import os

# Data from Excel
tahun = [2021, 2022, 2023, 2024, 2025]
produk = [683, 791, 546, 872, 925]
material = [1245, 1879, 1563, 1702, 1985]

# Create the figure and axis objects
fig, ax = plt.subplots(figsize=(10, 6))

# Plot data
ax.bar(np.array(tahun) - 0.2, produk, width=0.4, label='Jumlah Produk', color='blue')
ax.bar(np.array(tahun) + 0.2, material, width=0.4, label='Jumlah Material', color='green')

# Add titles and labels
ax.set_title('Perbandingan Jumlah Produk dan Material per Tahun', fontsize=15)
ax.set_xlabel('Tahun', fontsize=12)
ax.set_ylabel('Jumlah', fontsize=12)
ax.set_xticks(tahun)
ax.set_xticklabels(tahun)
ax.legend()
ax.grid(axis='y', linestyle='--', alpha=0.7)

# Save the figure
save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'produk_material_graph.png')
plt.savefig(save_path)
plt.close()
print(f'Graph saved to {save_path}')