import os
import shutil

def move_file(file_path, dest_folder):
    if os.path.isfile(file_path):
        file_size = os.path.getsize(file_path)
        if file_size < 20 * 1024: # 20 KB
            shutil.move(file_path, dest_folder)
            print(f"File {file_path} moved to {dest_folder}")
        else:
            print(f"File {file_path} is too big to move")
    else:
        print(f"{file_path} is not a valid file")

# Example usage
file_path = "path/to/file.txt"
dest_folder = "nobody"
move_file(file_path, dest_folder)