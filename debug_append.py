import pandas as pd
from pathlib import Path
from cleaner.engine import CleaningEngine

base = pd.DataFrame({
    'Order ID': ['A1'],
    'Nama': ['Budi'],
    'Harga': ['Rp 100.000'],
})
new = pd.DataFrame({
    'Order ID': ['A1', 'A3', 'A3', 'A2'],
    'Nama': ['Budi', 'Cici', 'Cici', 'Ayu'],
    'Harga': ['Rp 100.000', 'Rp 999.000', 'Rp 900.000', 'Rp 150.000'],
})
base_path = Path('tmp_append_debug.csv')
new_path = Path('tmp_append_debug_new.csv')
base.to_csv(base_path, index=False)
new.to_csv(new_path, index=False)
final, report = CleaningEngine().append_files(str(base_path), [str(new_path)], dedupe_keys=['Order ID'])
print(final)
print(report['total_dilewati'])
print(report['skipped_rows'])
