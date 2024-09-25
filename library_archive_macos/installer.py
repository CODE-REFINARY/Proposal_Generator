import os
import subprocess

# Get a list of all wheel files in the current directory
wheel_files = [file for file in os.listdir('.') if file.endswith('.whl')]

# Install each wheel file along with its dependencies
for wheel_file in wheel_files:
    subprocess.run(['pip', 'install', '--no-deps', wheel_file])