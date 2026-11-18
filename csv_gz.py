import gzip
import shutil
import os

# --------------------------------------------------
# 1. INPUT CSV FILE
# --------------------------------------------------

csv_path = input("Enter the full path of the CSV file: ").strip().strip('"')

# Check if file exists
if not os.path.isfile(csv_path):
    print("ERROR: File not found.")
    exit()

# --------------------------------------------------
# 2. OUTPUT FILE PATH
# --------------------------------------------------

gzip_path = csv_path + ".gz"

# --------------------------------------------------
# 3. COMPRESS CSV TO GZIP
# --------------------------------------------------

print("\nCompressing file...")

with open(csv_path, "rb") as f_in:
    with gzip.open(gzip_path, "wb") as f_out:
        shutil.copyfileobj(f_in, f_out)

# --------------------------------------------------
# 4. RESULT
# --------------------------------------------------

print("\nCompression completed successfully!")
print(f"Input file : {csv_path}")
print(f"Output file: {gzip_path}")

# File size comparison
original_size = os.path.getsize(csv_path)
compressed_size = os.path.getsize(gzip_path)

print(f"\nOriginal size   : {original_size / (1024**2):.2f} MB")
print(f"Compressed size : {compressed_size / (1024**2):.2f} MB")
print(f"Compression     : {(1 - compressed_size / original_size) * 100:.2f}%")