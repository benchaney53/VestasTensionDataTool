#import os
import sys

def fix_name(name: str) -> str:
    # Replace exactly three dashes with space-dash-space
    name = name.replace('---', ' - ')
    # Replace all other single dashes with a space
    name = name.replace('-', ' ')
    return name

def rename_in_tree(root_dir: str):
    for current_path, dirs, files in os.walk(root_dir, topdown=False):
        # Rename files
        for f in files:
            new_name = fix_name(f)
            if new_name != f:
                old_path = os.path.join(current_path, f)
                new_path = os.path.join(current_path, new_name)
                os.rename(old_path, new_path)
                print(f"Renamed file: {old_path} -> {new_path}")

        # Rename folders
        for d in dirs:
            new_name = fix_name(d)
            if new_name != d:
                old_path = os.path.join(current_path, d)
                new_path = os.path.join(current_path, new_name)
                os.rename(old_path, new_path)
                print(f"Renamed folder: {old_path} -> {new_path}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python rename_dashes.py <folder_path>")
        sys.exit(1)

    folder_path = sys.argv[1]
    rename_in_tree(folder_path)
