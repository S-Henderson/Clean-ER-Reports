import os

cur_dir = os.listdir(".")

for file_name in cur_dir:
    if file_name.startswith("Fraud Results"):
        os.rename(file_name, file_name.replace(" - CHECKED", ""))