********************************************************
𝓐𝓷𝓷𝓸𝓽𝓪𝓽𝓲𝓸𝓷 𝓘𝓷𝓽𝓮𝓻𝓶𝓮𝓻𝓰𝓮𝓻 - 𝓟𝓪𝓻𝓼𝓮𝓻 𝓟𝓻𝓸𝓰𝓻𝓪𝓶 (𝒜𝐼-𝒫𝒶𝓇𝓈𝑒𝓇)
********************************************************

Merge Roary, Prokka, eggNOG, KEGG and antiSMASH annotations into one TSV row per CDS.

This script is designed for an Excel workbook (.xlsx):
  - Output: optional pre-existing gene annotation table (header row present)
  - Roary: gene_presence_absence-style table (header row present)
  - Prokka: CDS table; supports the sample's headerless 7-column layout and a
            headered Prokka TSV-style table
  - eggnog: annotation table with Locus_Tag / Product / Gene columns
  - KEGG: either an ID-only, headerless list or a table with an ID column
  - antiSMASH: summary table with Region, Type, From, To and cluster columns

Dependencies:
    python -m pip install pandas openpyxl

Quick Example:
    python ai-parser.py --input "Sample data.xlsx" --out merged_gene_annotations.tsv

____________________________________________________________________________
INSTALLATION GUIDE:
Python Installation
2. Set up Python on Windows
1. Install Python 3 from the official Python downloads page.
2. During installation, check Add Python to PATH.
3. Open Command Prompt and verify Python is available:
   cmd     py --version     
   You should see a Python 3 version number. If py is unavailable, try python --version.
4. Create a folder for the parser, for example:
   text     C:\User\Documents\AI-Parser\     
   Put ai-parser-windows.py and your Excel workbook in this folder.
3. Run AI-Parser using Command Prompt
Suppose your folder contains:
C:\User\Documents\AI-Parser\
    ai-parser-windows.py
    Sample data.xlsx

USAGE:
1. Open Command Prompt and enter:
cd /d C:\Users\Documents\AI-Parser


Then run:
py -3 ai-parser.py --input "Sample data.xlsx" --out merged_gene_annotations.tsv

The parser will check its dependencies automatically before reading the workbook hehe.


BATCH-USAGE:
You can also use the batch launcher executable.
If you downloaded run-ai-parser.bat, keep it in the same folder as ai-parser-windows.py.
Run this from Command Prompt:
run-ai-parser.bat "Sample data.xlsx"

To specify an output filename:
run-ai-parser.bat "Sample data.xlsx" "final_annotations.tsv"



TROUBLESHOOTING:
Error:	What to do
1. 'py' is not recognized:	Install Python and enable PATH, or use python if available.
2. No module named pip:	Repair the Python installation or enable/install pip.
3. Dependency installation fails:	Check internet access, proxy settings, and package installation permissions.
4. Input workbook does not exist:	Check the filename and folder; retain quotation marks around filenames containing spaces.
5. Missing or incorrect annotations:	Verify the workbook sheet names and column headers match the supported layouts (see technical document).

____________________________________________________________________________
©2026 Zachary B. Lara