# 𝓐𝓷𝓷𝓸𝓽𝓪𝓽𝓲𝓸𝓷 𝓘𝓷𝓽𝓮𝓻𝓶𝓮𝓻𝓰𝓮𝓻 - 𝓟𝓪𝓻𝓼𝓮𝓻 𝓟𝓻𝓸𝓰𝓻𝓪𝓶 (𝒜𝐼-𝒫𝒶𝓇𝓈𝑒𝓇)

**AI-Parser** merges gene annotations from **Prokka, Roary, eggNOG, KEGG, and antiSMASH** into a single tab-separated values (TSV) file, with **one row per Prokka CDS**.

It is designed to help researchers consolidate gene identifiers, functional annotations, pangenome information, and biosynthetic gene cluster (BGC) context for downstream analysis.

## Features

- Uses Prokka CDS/locus tags as the main gene list when Prokka data are available.
- Matches Roary gene clusters using Prokka identifiers in isolate/GFF columns.
- Joins eggNOG annotations by query or locus tag, including normalization of common identifiers such as `P1_1_00959` → `P1_00959`.
- Checks whether each CDS identifier appears in the supplied KEGG list.
- Associates CDS coordinates with antiSMASH regions by coordinate overlap.
- Preserves an optional existing `Output` annotation sheet and fills missing fields where possible.
- Combines multiple matches into a single cell so the output remains one row per CDS.
- Writes a UTF-8 TSV file suitable for Excel, R, Python, and other downstream workflows.
- The Windows version checks for required Python packages and attempts to install missing dependencies.

## Input workbook

AI-Parser reads a single Excel workbook (`.xlsx`) containing worksheets with the following names (case-insensitive):

| Worksheet | Required? | Purpose |
|---|---|---|
| `Prokka` | Strongly recommended | Master list of CDS identifiers and genomic coordinates |
| `Roary` | Optional | Gene-cluster and isolate/presence metadata |
| `eggnog` | Optional | Functional annotation by query/locus tag |
| `KEGG` | Optional | CDS ID list and/or available KEGG annotation columns |
| `antiSMASH` | Optional | BGC region summary and coordinates |
| `Output` | Optional | Existing annotation table to preserve and supplement |

**The `Output` worksheet is not required.** It was used in the example workbook as an existing annotation table. If it is absent, AI-Parser creates the output columns and builds the gene list from Prokka. If Prokka has no usable records, the current script can fall back to the `Output` sheet; without either source, there is no master gene list to process.

For complete worksheet specifications, supported column headers, example layouts, and matching details, see **[`ai-parser_readme_supported_layout.md`](ai-parser_readme_supported_layout.md)**.

## Requirements

- Python 3
- `pandas`
- `openpyxl`

The Windows-enabled script checks whether `pandas` and `openpyxl` are installed and attempts to install any missing packages using `pip`. This requires a working Python installation and internet access when a package needs to be downloaded.

## Installation and usage

### Windows: Command Prompt

1. Install Python 3 from [python.org](https://www.python.org/downloads/windows/). During installation, enable **Add Python to PATH**.
2. Put `ai-parser.py`, `run-ai-parser.bat`, and your input workbook in a folder, for example:

   ```text
   C:\AI-Parser\
       ai-parser.py
       run-ai-parser.bat
       Sample data.xlsx
   ```

3. Open Command Prompt and navigate to the folder:

   ```bat
   cd /d C:\AI-Parser
   ```

4. Run the parser:

   ```bat
   py -3 ai-parser.py --input "Sample data.xlsx" --out merged_gene_annotations.tsv
   ```

If `py` is not recognized, try using `python` instead:

```bat
python ai-parser.py --input "Sample data.xlsx" --out merged_gene_annotations.tsv
```

### Windows: batch launcher

Keep `run-ai-parser.bat` in the same folder as `ai-parser.py`. Run:

```bat
run-ai-parser.bat "Sample data.xlsx"
```

To choose the output filename:

```bat
run-ai-parser.bat "Sample data.xlsx" "final_annotations.tsv"
```

Use quotation marks around filenames or paths that contain spaces.

### Manual dependency installation

If automatic installation fails, run:

```bat
py -3 -m pip install pandas openpyxl
```

Then run AI-Parser again.

## Output

The default output file is:

```text
merged_gene_annotations.tsv
```

The output contains one row per unique Prokka CDS identifier when Prokka records are available. It retains columns from an existing `Output` sheet, if provided, followed by derived annotation fields from the other sources.

Examples of merged fields include:

- Prokka locus tag, contig, feature, strand, coordinates, and product
- Roary gene cluster and annotation metadata
- eggNOG gene name, product, description, functional category, EC number, E-value, and bit score (when available)
- KEGG ID-list match status
- antiSMASH region, biosynthetic type, region coordinates, similarity confidence, and closest known cluster
- Match basis and CDS/region overlap status

Open the TSV in Excel and select **Tab** as the delimiter if prompted.

## Matching behavior and limitations

- **Prokka** supplies the preferred master CDS list.
- **Roary** matching depends on IDs in the relevant isolate/GFF columns corresponding to Prokka locus tags or normalized equivalents.
- **eggNOG** matching depends on a recognized query/locus-tag column and compatible IDs.
- **KEGG** ID-list matching indicates whether a CDS ID is present in the supplied list. If the KEGG worksheet contains only IDs, AI-Parser cannot infer KO numbers, pathways, or functional annotations that are not present in the source.
- **antiSMASH** matching is based on numeric coordinate overlap. The current implementation does not verify contig identity during this overlap check, so CDS and BGC coordinates must be from the same sequence/coordinate context. An overlap alone does not prove a gene's biological function in a BGC.
- Missing worksheets or unrecognized columns may result in blank annotations for that source.

## Troubleshooting

| Problem | Suggested check |
|---|---|
| `'py' is not recognized` | Install Python and enable PATH, or try `python` instead. |
| `No module named pip` | Repair the Python installation or enable/install `pip`. |
| Dependency installation fails | Check internet access, proxy settings, and package permissions. |
| Input workbook not found | Confirm the path and filename; quote paths containing spaces. |
| Roary fields are blank | Confirm the GFF/isolate column IDs match Prokka locus tags. |
| eggNOG fields are blank | Confirm the sheet has a recognized query/locus-tag column and matching IDs. |
| antiSMASH fields are blank | Confirm the region summary has numeric start/end coordinates compatible with Prokka coordinates. |
| KEGG has matches but no pathway details | Check whether the KEGG worksheet contains annotation columns; an ID-only list provides only ID matching. |

## Project files

| File | Description |
|---|---|
| `ai-parser.py` | Main parser |
| `run-ai-parser.bat` | Optional Windows launcher |
| `ai-parser_readme.md` | Original short project notes |
| `ai-parser_readme_supported_layout.md` | Detailed workbook layout and matching guide |
| `README.md` | Main GitHub project page |

## License

Add a `LICENSE` file to the repository to make the reuse and redistribution terms explicit. If you choose the MIT License, use the official [MIT License text](https://opensource.org/license/mit) and confirm that you have the right to license the project under your name.

---

Copyright © 2026 Zachary B. Lara
