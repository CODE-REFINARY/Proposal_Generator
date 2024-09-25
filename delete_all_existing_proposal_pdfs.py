import os
from decouple import config

def delete_files_by_pattern(root_dir, dry_run=False):
    # Iterate through all directories and subdirectories
    for dir_path, _, files in os.walk(root_dir):
        # Search for files matching the pattern in the current directory
        for file_name in files:
            # Check if the file name matches the pattern
            if (len(file_name) == 10 and file_name[:6].isdigit() and file_name[6] == '.' and file_name.endswith('.pdf')) or (file_name.startswith("TEMP")) or (file_name.startswith("SITE")):
                file_to_delete = os.path.join(dir_path, file_name)
                if dry_run:
                    print(f"Would delete: {file_to_delete}")
                else:
                    try:
                        os.remove(file_to_delete)
                        print(f"Deleted: {file_to_delete}")
                    except Exception as e:
                        print(f"Error deleting {file_to_delete}: {e}")


# Specify the root directory where you want to start searching
root_directory = config("PROPOSALS_BASE_DIR")

# Call the function to delete files with the specified pattern
delete_files_by_pattern(root_directory, True)
