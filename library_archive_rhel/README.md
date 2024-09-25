# Library Archive Directory for RHEL

This directory contains wheel files for all of the libraries that the PDF Generator tool uses. This also includes libraries that the LibreOffice python install needs. To install these files simply run the installer.py file. It's recommended to create a venv before doing this. Once the venv is activated this script should be run with any version of python (I don't think it should matter) and the libraries will be installed for that venv. This directory also contains a copy of the get-pip.py file which can be used to install pip completely offline.

Note: all wheels are only tested to work with Centos 7. Wheel's are not necessarily OS-independent.
