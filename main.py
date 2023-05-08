import os
import shutil
import glob

# Set the source and destination directories
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
destination_folder = "C:\\excel"

# Create the destination folder if it does not exist
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# Find Excel files in the Downloads folder (both .xls and .xlsx)
excel_files = glob.glob(os.path.join(downloads_folder, "*.xls")) + glob.glob(os.path.join(downloads_folder, "*.xlsx"))

# Move the Excel files to the destination folder
for excel_file in excel_files:
    try:
        shutil.move(excel_file, os.path.join(destination_folder, os.path.basename(excel_file)))
        print(f"Moved {excel_file} to {destination_folder}")
    except shutil.Error as e:
        print(f"Error moving {excel_file}: {e}")

print("All Excel files have been moved.")
