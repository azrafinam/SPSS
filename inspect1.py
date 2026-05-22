import pyreadstat
import pandas as pd

df, meta = pyreadstat.read_sav("data.sav")

# show ALL column names properly
print("\n===== COLUMN NAMES =====\n")
for col in df.columns:
    print(col)

print("\n===== DATA SHAPE =====")
print(df.shape)

print("\n===== FIRST 5 ROWS =====\n")
print(df.head().to_string())