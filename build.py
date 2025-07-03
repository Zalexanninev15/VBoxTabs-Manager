"""
This script automates the local build process for the VBoxTabs-Manager project.

It performs the following steps:
1.  Cleans up previous build artifacts.
2.  Installs required Python dependencies from requirements.txt, plus nuitka and zstandard.
3.  Extracts the version number from VBoxTabs-Manager.py.
4.  Builds a standalone executable using Nuitka.
5.  Renames the output directory and creates a distributable ZIP archive.

To run this script:
- Ensure you have Python 3.13+ installed.
- Open a Windows Terminal (or CMD/PowerShell) in the project root directory.
- Run the command: python build.py
"""

import os
import re
import shutil
import subprocess
import sys

# --- Configuration ---
SOURCE_FILE = "VBoxTabs-Manager.py"
REQUIREMENTS_FILE = "requirements.txt"
DIST_DIR_ORIGINAL = "VBoxTabs-Manager.dist"
DIST_DIR_RENAMED = "VBoxTabs Manager"
ZIP_FILENAME = "VBoxTabs-Manager" # .zip will be added by shutil

def check_python_version():
    """Checks if the current Python version is 3.13 or newer."""
    print("--- Verifying Python version ---")
    if sys.version_info < (3, 13):
        print(f"Error: This script requires Python 3.13 or newer. You are using {sys.version}")
        sys.exit(1)
    print(f"Python version {sys.version} is compatible.")

def clean_previous_builds():
    """Removes artifacts from previous builds to ensure a clean slate."""
    print("\n--- Cleaning up previous build artifacts ---")
    artifacts = [
        DIST_DIR_ORIGINAL,
        DIST_DIR_RENAMED,
        f"{ZIP_FILENAME}.zip",
        "VBoxTabs-Manager.build", # Nuitka build folder
        "VBoxTabs Manager.exe"    # Nuitka single-file output
    ]
    for artifact in artifacts:
        if os.path.isfile(artifact):
            print(f"Removing file: {artifact}")
            os.remove(artifact)
        elif os.path.isdir(artifact):
            print(f"Removing directory: {artifact}")
            shutil.rmtree(artifact)
    print("Cleanup complete.")

def install_dependencies():
    """Installs dependencies from requirements.txt and other necessary packages."""
    print("\n--- Installing dependencies ---")
    try:
        # Use sys.executable to ensure we use the pip from the correct python env
        python_executable = sys.executable
        
        print(f"Installing packages from {REQUIREMENTS_FILE}...")
        subprocess.run([python_executable, "-m", "pip", "install", "-r", REQUIREMENTS_FILE], check=True)
        
        print("Installing Nuitka and Zstandard...")
        subprocess.run([python_executable, "-m", "pip", "install", "nuitka", "zstandard"], check=True)
        
        print("Dependencies installed successfully.")
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"Error during dependency installation: {e}")
        sys.exit(1)

def get_version_from_source():
    """Extracts the version string from the main Python source file."""
    print("\n--- Extracting version from VBoxTabs-Manager.py ---")
    version_pattern = re.compile(r'version_label\s*=\s*QLabel\("Version ([0-9.]+[a-zA-Z0-9_-]*)"\)')
    try:
        with open(SOURCE_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
            match = version_pattern.search(content)
            if match:
                version = match.group(1)
                print(f"Version detected: {version}")
                return version
            else:
                print("Warning: No version found in source file. Using default '0.0'.")
                return "0.0"
    except FileNotFoundError:
        print(f"Error: Source file '{SOURCE_FILE}' not found.")
        sys.exit(1)

def run_nuitka_build(version):
    """Compiles the Python script into a standalone executable using Nuitka."""
    print("\n--- Building executable with Nuitka ---")
    
    # Nuitka command arguments, mirroring the YAML file
    nuitka_args = [
        sys.executable, "-m", "nuitka",
        "--assume-yes-for-downloads",
        "--standalone",
        "--follow-imports",
        "--enable-plugin=pyside6",
        "--include-package=win32gui",
        "--include-package=win32process",
        "--include-package=win32con",
        "--include-package=win32api",
        "--include-package=qdarkstyle",
        "--static-libpython=no",
        "--company-name=Zalexanninev15",
        "--product-name=VBoxTabs Manager",
        f"--file-version={version}",
        "--lto=yes",
        "--show-memory",
        "--show-progress",
        "--nofollow-import-to=tkinter",
        "--nofollow-import-to=PIL",
        "--nofollow-import-to=numpy",
        "--jobs=8",
        "--windows-console-mode=disable",
        "--remove-output",
        "-o", "VBoxTabs Manager.exe",
        SOURCE_FILE
    ]
    
    print(f"Running Nuitka with version {version}...")
    try:
        subprocess.run(nuitka_args, check=True)
        print("Nuitka build completed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error during Nuitka build: {e}")
        sys.exit(1)

def create_zip_archive():
    """Renames the build output directory and creates a ZIP archive."""
    print("\n--- Creating ZIP archive ---")
    if not os.path.isdir(DIST_DIR_ORIGINAL):
        print(f"Error: Nuitka output directory '{DIST_DIR_ORIGINAL}' not found.")
        sys.exit(1)
    
    # 1. Rename the directory
    print(f"Renaming '{DIST_DIR_ORIGINAL}' to '{DIST_DIR_RENAMED}'")
    os.rename(DIST_DIR_ORIGINAL, DIST_DIR_RENAMED)
    
    # 2. Create the ZIP archive
    print(f"Creating '{ZIP_FILENAME}.zip'...")
    shutil.make_archive(
        base_name=ZIP_FILENAME,
        format='zip',
        root_dir='.',
        base_dir=DIST_DIR_RENAMED
    )
    print("ZIP archive created successfully.")

def main():
    """Main function to run the build process."""
    print("Starting local build process for VBoxTabs Manager...")
    
    check_python_version()
    clean_previous_builds()
    install_dependencies()
    version = get_version_from_source()
    run_nuitka_build(version)
    create_zip_archive()
    shutil.rmtree(DIST_DIR_RENAMED)
    
    print("\n" + "="*50)
    print("✅ BUILD SUCCEEDED!")
    print("="*50)
    print(f"Build artifacts are ready:")
    print(f"  - ZIP File:  ./{ZIP_FILENAME}.zip")
    print("\n--- Manual Release Steps ---")
    print("The automated GitHub Release step from the workflow is not run locally.")
    print("To create a new release on GitHub:")
    print("1. Go to your repository's 'Releases' page on GitHub.")
    print("2. Click 'Draft a new release'.")
    print(f"3. Use 'v{version}' as the tag name.")
    print(f"4. Use 'Version {version}' as the release title.")
    print("5. Write your release notes (you can copy the latest commit message).")
    print(f"6. Upload the '{ZIP_FILENAME}.zip' file as a release asset.")
    print("7. Publish the release.")

if __name__ == "__main__":
    main()