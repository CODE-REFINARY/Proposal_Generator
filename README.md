# Word Document Templating System
PDF Generator is an implementation of the previous webform-to-PDF converter approach that uses MS Word documents and Python instead of just PHP. The high level approach is as follows:

Define directory with MS Word document templates that each are written with identifiers following a templating scheme.

Run a python script from the command line that imports python-docx (https://python-docx.readthedocs.io/en/latest/).

Use python-docx to search for every occurence of a particular string in the Word document (template tag) and replace that it with a database entry. Do this once for every entry in a pre-defined list of template tags that are expected to appear in the Word document. Do the same for conditional template tags.

Save all the resultant edited Word documents as new Word documents and then merge them into a single large Word document

Call LibreOffice from the command line (using the python subprocess command) to convert the large Word document to PDF.

Merge user upload PDFs into the PDF.

Insert bookmarks into the PDF by extracting text from each page and matching it with expected text to determine which pages to bookmark on.

# Installing a Copy of Python 3.9 on PDBDEV/PDBPRO
Both servers are currently on Python 3.6. Psycopg2 (one of our dependencies) requires at least Python 3.7. Thus we will install a more up-to-date copy of python in /usr/local

Run the following commands on the CentOS 7 server to accomplish this:

# Start by making sure your system is up-to-date:
yum update
# Compilers and related tools:
yum groupinstall -y "development tools"
# Libraries needed during compilation to enable all features of Python:
yum install -y zlib-devel bzip2-devel openssl-devel ncurses-devel sqlite-devel readline-devel tk-devel gdbm-devel db4-devel libpcap-devel xz-devel expat-devel

# Python 3.9.18:
wget http://python.org/ftp/python/3.9.18/Python-3.9.18.tar.xz
tar xf Python-3.9.18.tar.xz
cd Python-3.9.18
./configure --prefix=/usr/local --enable-shared LDFLAGS="-Wl,-rpath /usr/local/lib"
make && make altinstall

# Strip the Python 3.9 binary (this should reduce the size of the binary to about 1/4):
strip /usr/local/lib/libpython3.9.so.1.0

# Pip should be installed automatically in Python 3.9 but in case it isn't:
wget https://bootstrap.pypa.io/get-pip.py
# Then execute it:
python3.9 get-pip.py
LibreOffice Installation
LibreOffice is the program that converts Word documents to PDF - a crucial aspect of this PDF Generator tool. LibreOffice is typically used as a GUI application but can also be used in headless mode which is how we use it (from the command line). LibreOffice has been installed on pdbvdev.org. The path to the launch script is /localdisk/apps/LibreOffice/opt/libreoffice7.6/program/soffice. The command to convert input.docx to a pdf in the current directory is /localdisk/apps/LibreOffice/opt/libreoffice7.6/program/soffice --convert-to pdf input.docx --outdir .

It is this command that is called by the PDF Generator script to accomplish the actual Word → PDF conversion.

To download LibreOffice go to the official website and download version 7.6 for whatever version is appropriate for the target system.

The Pdbdev LibreOffice installation is at /localdisk/apps/LibreOffice/opt/libreoffice7.6/program

Daemonizing Unosever
The --daemon flag for unoserver doesn’t work so to daemonize this process we use a python library and modify the python unoserver library code a tiny bit. First ensure that unoserver is installed. Important: Unoserver must be installed for the python installation that LibreOffice uses. For pdbdev this installation is located at /localdisk/apps/LibreOffice/opt/libreoffice7.6/program/python. The wheel for unosever is located in the library_archive folder along with all the others wheels. To install unosever for the python that LibreOffice uses one must first ensure that this python has pip installed. This can be done by running the following inside of library_archive:

/absolute/path/to/python get-pip.py

/absolute/path/to/python installer.py

Next, navigate to the library code for the unoserver library. On pdbdev this is at

/localdisk/apps/LibreOffice/opt/libreoffice7.6/program/python-core-3.8.18/lib/python3.8/site-packages/unoserver

Then, open up the server.py file and scroll down to the bottom. Replace this snippet:

if __name__ == "__main__":
    main()
…with this snippet:

if __name__ == "__main__":
    with daemon.DaemonContext():
        main()
With this modification made the daemon process can be launched with the following command:

/localdisk/apps/LibreOffice/opt/libreoffice7.6/program/python -m unoserver.server --executable /localdisk/apps/LibreOffice/opt/libreoffice7.6/program/soffice

When the daemon process is running conversions are initiated with the following command:

/localdisk/apps/LibreOffice/opt/libreoffice7.6/program/python -c 'import unoserver.client as t; t.converter_main()' /abs/path/to/input.docx /abs/path/to/output.pdf

IMPORTANT make sure that the daemon process is set as a service daemon so that it runs on system start.

Installing Python Libraries
The python libraries needed for this application can be installed by first navigating to the library_archive folder in the top level directory. This folder contains wheels for all the libraries to be installed. To install these libraries you can run installer.py (also located in this directory) with the python install that you want the libraries to be installed for (i.e. /abs/path/to/python3.9 installer.py). This script just installs all of the wheels in this directory. It might be best to first create a virtual env before running this script if this python install is used for other applications on the system too.

Deprecated Items
All notes below are for previous implementations of the PDF converter that were not pursued but are kept for reference:

REPORTLAB
ReportLab Open Source Installation
ReportLab is a PDF generation library written in Python. It lets you create pdfs from scratch and comes in two flavors: a free open source version that’s relatively low-level and advanced and a “pro” a yearly paid version that bundles a templating engine alongside the generation engine (which comes with the open source version) to streamline the process of specifying the look and feel of the output PDFs.

At the time of writing (October 2023) the current version of ReportLab is 4.0.6. Install ReportLab by running:

pip install reportlab==4.0.6

In order to create rich PDFs with bitmap and vector images Cairo must also be installed:

pip install pycairo==1.25.1

If you encounter:ERROR: Failed building wheel for pycairo then run the following and then try installing pycairo again (first install homebrew if you haven’t already):

brew install cairo

brew install pkg-config

pypdf is a free and open-source pure-python PDF library capable of splitting, merging, cropping, and transforming the pages of PDF files. It can be used on PDFs created by ReportLab and will be used to create bookmarks in the pdf and merge pdfs together. pypdf is the sequel to the famous PyPDF2 library (and is currently maintained by the same developer). Install it by running the following:

pip install pypdf==3.16.4

ReportLab Plus Installation
Follow these steps: https://docs.reportlab.com/install/ReportLab_Plus_version_installation/

High Level PDF Generator Implementation
A python file called gen_proposal.py is called from a shell with a proposal.json file as a first argument and the proposal number as the second argument.

gen_proposal.py creates a python subprocess that runs php extract_constants.php as a shell command.

This command extracts the constant value definitions in common/pdb_common.php and writes them as JSON to STDOut which is read by gen_proposal.py into a python dictionary object, config.

The values in config are used to create a database connection and specify any other flags and relevant directories that will be used in the PDF generation process

gen_proposal.py then queries the database view psycopg2 and pulls records corresponding to the proposal number command line argument, storing the data in a series of python variables

gen_proposal.py then creates an RML template string from the first command line argument, proposal.json which is loaded into a jsondict.p module dictionary object. The python variables containing field values from the database are also inserted into the dictionary prior to the dictionary being synthesized into an RML bytes string via getOutput.

The RML string is loaded into a call to rml2pdf.go() along with an output directory and the PDF is generated.

Bookmark creation and concatenation details to follow…

Headless Browser Print to PDF System
High Level Implementation
Create a YII action that expects an HTTP header with the Proposal ID. When called it returns a page that looks like a PDF template with the web form fields filled in with database values from the specified Proposal. This action is only returned if the caller matches the server’s IP address (nobody but the server should be able to access this page).

Run headless chrome in a command line process like so chrome --headless --print-to-pdf https://url where url goes to the YII action defined above. This command will print the page and save it to a PDF (server-side rendering).

This PDF is now ready to bookmarked and merged.


## LibreOffice Daemon Process Command on Local Machine
<code>
/Applications/LibreOffice.app/Contents/Resources/python -m unoserver.server --executable /Applications/LibreOffice.app/Contents/MacOS/soffice
</code>

## LibreOffice Conversion Command on Local Machine
<code>
/Applications/LibreOffice.app/Contents/Resources/python -c 'import unoserver.client as t; t.converter_main()' /Users/gdubinin/Documents/python-docx/result.docx /Users/gdubinin/Documents/python-docx/output.pdf
</code>