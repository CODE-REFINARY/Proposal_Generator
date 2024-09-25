from decouple import config
from subprocess import run
import subprocess
from pdf_gen_helper_functions import validate_path
import argparse
import psycopg2
import time
import sys
from os.path import join

# Start the timer which is used to print out how long the PDF generation process takes.
tic = time.perf_counter()

# Extract the proposal ID command line arg to determine which proposal to generate a PDF for.
parser = argparse.ArgumentParser(description="Generate a PDF given a proposal ID")
parser.add_argument("proposal_id", type=str, help="This str specifies the proposal id to generate.")
parser.add_argument("--coversheet-only", "-c", action="store_true", help="This bool tells the program to produce only the coversheet.")
parser.add_argument("--output-filename", "-o", action="store", help="This argument should be a string which will be the name of the output PDF.")
args = parser.parse_args()
PROPOSAL_ID = args.proposal_id

# Get the directory for the proposal now that we know what it is
PROPOSALS_BASE_DIR = validate_path(config("PROPOSALS_BASE_DIR"))
PROPOSAL_DIR = validate_path(join(PROPOSALS_BASE_DIR, PROPOSAL_ID))

# Make sure the directory for storing log messages exists (and create one if it doesn't)
PDF_LOGS_DIR = validate_path(join(PROPOSAL_DIR, "pdf_logs"))
PDF_LOG_FILE = "pdf_gen_logs.txt"

# Determine what p_type of proposal this ID corresponds to so that we can call the appropriate controller
# function to generate a PDF for it. Read the proposal p_type from the database so that we know which controller
# to call later.
conn = psycopg2.connect(  # This variable links the iodpdatadev database to this program
    host=config("DB_HOST"),
    database=config("DB_DATABASE"),
    user=config("DB_USERNAME"),
    password=config("DB_PASSWORD"),
    port=config("DB_PORT")
)
cur = conn.cursor()
cur.execute("SELECT * FROM proposal WHERE id = %s", (PROPOSAL_ID,))
PROPOSAL_ROW = cur.fetchone()
PROPOSAL_COLS = [desc[0] for desc in cur.description]

DRILLING_TYPES = ["Full", "APL", "Pre", "Add", "CPP", "SRR"]
LEAP_TYPES = ["Pre-LEAP", "Full-LEAP"]

if not PROPOSAL_ROW:
    raise RuntimeError("This proposal does not yet exist in the database.")

proposal_type = PROPOSAL_ROW[PROPOSAL_COLS.index("proposal_type")]
c_only = args.coversheet_only
output_name = args.output_filename

# Decide which controller to call based on the newly determined proposal_type.
if proposal_type in DRILLING_TYPES:
    cmd = [config("PRIMARY_PYTHON_INSTALLATION_PATH"), config("PDF_GEN_DIR") + "controller_drilling.py", PROPOSAL_ID]
elif proposal_type in LEAP_TYPES:
    cmd = [config("PRIMARY_PYTHON_INSTALLATION_PATH"), config("PDF_GEN_DIR") + "controller_leaps.py", PROPOSAL_ID]
else:
    raise ValueError("This proposal type is not yet supported for PDF generation.")

if c_only:
    cmd.append("--coversheet-only")
if output_name:
    cmd.append("--output-filename")
    cmd.append(output_name)

with open(join(PDF_LOGS_DIR, PDF_LOG_FILE), "w") as log:
    print("Running the following command and writing output to " + join(PDF_LOGS_DIR, PDF_LOG_FILE))
    print(" ".join(cmd))
    proc = run(cmd, stdout=log, stderr=subprocess.STDOUT, text=True)

toc = time.perf_counter()
print(f"Completed in {toc - tic:0.4f} seconds")

if proc.returncode == 0:
    print("Operation Successful")
    sys.exit(0)
else:
    print("Operation Failed")
    sys.exit(1)
