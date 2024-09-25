#!/bin/bash

# Check if a number is provided as a parameter
if [ $# -eq 0 ]; then
    echo "Usage: $0 <number>"
    exit 1
fi

# Assign the provided number to a variable
number=$1
password="piquing.classy.victory"
remote_command="/localdisk/apps/pdf_generator/myvenv/bin/python3.11 /localdisk/apps/pdf_generator/gen.py ${number}"

# Specify the server details
server="iodp@iodpdev.iodp.org"
remote_file="/localdisk/data/pdbdev_files/${number}/${number}.pdf"

sshpass -p "$password" ssh $server "$remote_command"

# Copy the file using scp
sshpass -p "$password" scp $server:$remote_file ./${number}.pdf

# Check if scp was successful
if [ $? -eq 0 ]; then
    echo "File successfully copied."

    # Open the resulting file
    if command -v xdg-open &> /dev/null; then
        xdg-open ${number}.pdf
    elif command -v open &> /dev/null; then
        open ${number}.pdf
    else
        echo "Cannot open file. Please open result_file.txt manually."
    fi
else
    echo "Failed to copy the file. Check your parameters and server configuration."
fi

