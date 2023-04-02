import os
import time
import webbrowser
import shutil
from datetime import datetime

urls_file = os.path.join(r'C:\Users\e400284\Xperi\Premkumar Jaganathan (CTR) - SDL Report\kgvt-report\urls.txt')

# read the URLs from the file and store them in a list
with open(urls_file, 'r') as file:
    urls = file.readlines()
    urls = [url.strip() for url in urls]  # remove any whitespace or newlines

download_path = os.path.expanduser('~\Downloads')
destination_folder = os.path.join(r'C:\Users\e400284\Xperi\Premkumar Jaganathan (CTR) - SDL Report\kgvt-report\reports')
destination_folder2 = os.path.join(r'C:\Users\e400284\Xperi\Premkumar Jaganathan (CTR) - SDL Report\kgvt-report\ingest-files')

# get the current list of files in the download path
original_files = set(os.listdir(download_path))

# open the download urls in new tabs
for url in urls:
    webbrowser.open_new_tab(url)
    time.sleep(5)  # wait for the download to start before closing the tab

# wait until the new files have finished downloading
while True:
    # get the updated list of files in the download path
    new_files = set(os.listdir(download_path))
    # check if any new files have been added to the download path
    if new_files != original_files:
        # get the downloaded xlsx files
        downloaded_files = [f for f in new_files - original_files if f.endswith('.xlsx')]
        failed_files = [f for f in new_files - original_files if not f.endswith('.crdownload') and f not in downloaded_files]
        break
    # wait for 5 seconds before checking again
    time.sleep(5)

# create the folder with today's date in the second destination
today_folder = os.path.join(destination_folder2, datetime.today().strftime('%d-%b'))
os.makedirs(today_folder, exist_ok=True)

# move the downloaded xlsx files to the first destination folder
num_files_moved = 0
for file in downloaded_files:
    file_path = os.path.join(download_path, file)
    dest_path = os.path.join(destination_folder, file)
    shutil.copy(file_path, destination_folder2)
    shutil.move(file_path, dest_path)
    num_files_moved += 1

print(f"Moved {num_files_moved} files to {destination_folder} and copied to {destination_folder2}")

# move the copied xlsx files to the today's folder in the second destination
copied_files = [f for f in os.listdir(destination_folder2) if f.endswith('.xlsx')]
num_files_moved = 0
for file in copied_files:
    file_path = os.path.join(destination_folder2, file)
    dest_path = os.path.join(today_folder, file)
    shutil.move(file_path, dest_path)
    num_files_moved += 1

print(f"Moved {num_files_moved} files to {today_folder}")

if failed_files:
    print(f"The following files failed to download: {failed_files}")
else:
    print("All files downloaded successfully.")
