import os
import shutil
import zipfile
from pathlib import Path
import fnmatch

def create_archive(output_filename, input_files, exclude_files=None):
    """
    Create a zip archive from a list of input files.

    Args:
        output_filename (str): Path to the output zip file.
        input_files (list): List of file paths or glob patterns to include in the archive.
        exclude_files (list): List of file names or patterns to exclude from the archive.
    """
    exclude_files = exclude_files or []

    with zipfile.ZipFile(output_filename, 'w') as archive:
        for pattern in input_files:
            for file in Path().glob(pattern):
                if file.is_file() and not any(fnmatch.fnmatch(file.name, excl) for excl in exclude_files):
                    archive.write(file, arcname=file.name)


def copy_files(source_dir, destination_dir):
    """
    Copy all files from the source directory to the destination directory.

    Args:
        source_dir (str): Path to the source directory.
        destination_dir (str): Path to the destination directory.
    """
    source_path = Path(source_dir)
    dest_path = Path(destination_dir)

    if not source_path.is_dir():
        raise ValueError(f"Source directory {source_dir} does not exist.")

    # Create the destination directory if it doesn't exist
    dest_path.mkdir(parents=True, exist_ok=True)

    for item in source_path.iterdir():
        if item.is_file():
            shutil.copy2(item, dest_path / item.name)
        elif item.is_dir():
            # Recursively copy subdirectories
            shutil.copytree(item, dest_path / item.name, dirs_exist_ok=True)


def create_directory(directory_path):
    """
    Create a new directory if it doesn't exist.

    Args:
        directory_path (str): Path of the directory to create.
    """
    # Use makedirs with exist_ok=True to avoid error if directory exists
    os.makedirs(directory_path, exist_ok=True)


def delete_files_and_directories(targets, exclude_files=None):
    """
    Delete files and directories matching the provided patterns, excluding specified files or patterns.

    Args:
        targets (list): List of file paths or glob patterns to delete.
        exclude_files (list): List of file paths or glob patterns to exclude from deletion.
    """
    exclude_files = exclude_files or []


    for pattern in targets:
        for item in Path().glob(pattern):
            # Resolve the full path for accurate comparison with exclude patterns
            item_path = str(item.resolve())
            if not any(fnmatch.fnmatch(item_path, str(Path(excl).resolve())) for excl in exclude_files):
                if item.is_file():
                    item.unlink()  # Delete the file
                elif item.is_dir():
                    shutil.rmtree(item)  # Delete the directory

def copy_files_and_directories(source_dir, destination_dir, inclusion_list=None, exclusion_list=None):
    """
    Copies files from a source directory to a destination directory based on inclusion and exclusion lists.

    :param source_dir: The source directory path.
    :param destination_dir: The destination directory path.
    :param inclusion_list: List of patterns for files to include. If provided, only these files are copied.
    :param exclusion_list: List of patterns for files to exclude. If provided, these files are not copied.
    """
    if not os.path.exists(destination_dir):
        os.makedirs(destination_dir)

    # Get list of all files in source directory
    for root, _, files in os.walk(source_dir):
        matching_files = []
        for file_name in files:
            # Check inclusion and exclusion lists
            if inclusion_list and not any(fnmatch.fnmatch(file_name, pattern) for pattern in inclusion_list):
                continue
            if exclusion_list and any(fnmatch.fnmatch(file_name, pattern) for pattern in exclusion_list):
                continue
            matching_files.append(file_name)

        if matching_files:
            relative_path = os.path.relpath(root, source_dir)
            dest_subdir = os.path.join(destination_dir, relative_path)
            if not os.path.exists(dest_subdir):
                os.makedirs(dest_subdir)

            # Copy only matching files
            for file_name in matching_files:
                source_file_path = os.path.join(root, file_name)
                destination_file_path = os.path.join(dest_subdir, file_name)
                shutil.copy2(source_file_path, destination_file_path)


def delete_directory(directory_path):
    """
    Deletes a directory and all its contents.

    :param directory_path: The path to the directory to delete.
    """
    try:
        shutil.rmtree(directory_path)
    except Exception as e:
        print(f"Error: Unable to delete directory '{directory_path}'. {e}")

def insert_string_before_extension(directory, insert_string):
    """
    Inserts a given string before the file extension for all files in the given directory.

    Args:
        directory (str): Path to the directory containing the files.
        insert_string (str): The string to insert before the file extensions.
    """
    if not os.path.isdir(directory):
        print(f"Error: {directory} is not a valid directory.")
        return

    for filename in os.listdir(directory):
        # Split filename into name and extension
        name, ext = os.path.splitext(filename)
        if ext:  # Only process files with extensions
            new_name = f"{name}{insert_string}{ext}"
            old_path = os.path.join(directory, filename)
            new_path = os.path.join(directory, new_name)
            os.rename(old_path, new_path)