import os
import pwd
from typer import Typer
from tqdm.auto import tqdm
from collections import defaultdict
import time
import openpyxl

app = Typer()

@app.command()
def find_large_files(path: str = ".", minimum_size: float = 10.0, unit: str = "MB") -> None:
    #I have Entered Defaults as the Current Working Directory, 10 MB and MB as the Unit

    user_storage = defaultdict(int)

    #Valid Units for the Sizes
    valid_units = ["MB", "KB", "GB"]

    #Check if the Unit Entered is Valid
    try:
        assert unit in valid_units
    except ValueError as e:
        print(f"Unit Entered is Invalid: {e}. Please Enter a Valid Unit from {valid_units}")


    multiplier = {"MB": 1024**2, "KB": 1024, "GB": 1024**3}[unit]
    size_threshold_bytes = minimum_size * multiplier

    for root, _, files in os.walk(path):
        for file in files:
            file_path = os.path.join(root, file)
            file_stat = os.stat(file_path)
            file_size = file_stat.st_size

    Files_Found = []

    total_files = sum(len(files) for _, _, files in os.walk(path))  # Count total files

    #A Fancy Progress Bar for Searching Files
    with tqdm(total=total_files, desc="Searching for large files...") as pbar:
        for root, _, files in os.walk(path):
            for file in files:
                file_path = os.path.join(root, file)
                file_stat = os.stat(file_path)
                file_size = file_stat.st_size
                time.sleep(0.01)
                pbar.update()

                #Main If Condition for the File Size Greater than Minimum Specified Size
                if file_size > size_threshold_bytes:
                    #Check if the User ID can be fetched
                    try:
                        user_id = file_stat.st_uid
                        user_name = pwd.getpwuid(user_id).pw_name
                        user_storage[user_name] += file_size
                    except Exception as e:
                        user_name = "Unavailable"

                    formatted_size = f"{file_size / multiplier:.2f}{unit}"
                    File_Found = [file_path, formatted_size, user_name]
                    Files_Found.append(File_Found)

    for File in Files_Found :
        print(f"{File[0]}: {File[1]} (Created by: {File[2]})")

    top_users = sorted(user_storage.items(), key=lambda item: item[1], reverse=True)

    #Open Excel Workbook
    wb = openpyxl.Workbook()
    ws = wb.active

    #Title
    ws.title = "Top Users by Storage Usage"

    #Headers
    ws.cell(row=1, column=1).value = "User"
    ws.cell(row=1, column=2).value = f"Storage ({unit})"

    #Top 5 Users
    for row, (user, storage) in enumerate(top_users, start=2):
        formatted_storage = f"{storage / multiplier:.2f}"
        ws.cell(row=row, column=1).value = user
        ws.cell(row=row, column=2).value = formatted_storage

    # Save Excel
    wb.save("Top_Users.xlsx")

    print("\nTop Users Data is Exported")


if __name__ == "__main__":
    app()
