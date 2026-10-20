# AI-Parser — Annotation Intermerger

**Version: `0.2.0`** (development version)

The GitHub version `0.1.2` remains unchanged until you publish a new release. This local `0.2.0` version adds separate per-tool inputs and Bakta support. Check the installed script version with `py -3 ai-parser.py --version`.

AI-Parser merges gene/CDS annotations from **Bakta or Prokka**, **Roary**, **eggNOG-mapper**, **KEGG/BlastKOALA**, and **antiSMASH** into one UTF-8 tab-separated file, with one row per unique master CDS identifier.

## Versioning

AI-Parser uses `MAJOR.MINOR.PATCH` versioning. While the project is in `0.x`, increment the minor number for a meaningful feature release and the patch number for bug fixes. This working copy identifies itself as `0.2.0`; the existing GitHub release can remain `0.1.2` until you commit and publish the newer version.

## What changed

- Each tool can now be supplied as its own independent file; no combined Excel workbook is required.
- **Bakta and Prokka are alternative master gene lists.** Bakta takes precedence if both are supplied.
- Roary, eggNOG, KEGG/BlastKOALA, and antiSMASH are all optional.
- An Excel file (`.xlsx`/`.xlsm`) can be supplied for each source instead of its usual text output.
- The parser uses file extension and column headers/content to read supported formats. If a table's headers do not identify it confidently, the program reports an error with the observed headers instead of silently treating it as the wrong tool.
- The original multi-sheet workbook mode remains available with `--input` for backwards compatibility.

## Supported inputs

| Tool | Preferred native output | Other supported inputs | Notes |
|---|---|---|---|
| **Bakta** | `*.tsv` | `.csv`, `.xlsx`/`.xlsm`, `.gff`/`.gff3` | Use the main Bakta TSV or GFF3 with locus tags and coordinates. Bakta is the master list if supplied. |
| **Prokka** | `*.tsv` | `.csv`, `.xlsx`/`.xlsm`, `.gff`/`.gff3` | Prokka TSV and common headered feature tables are supported; legacy headerless 7-column extracts are supported in workbook mode. |
| **Roary** | `gene_presence_absence.csv` | `.tsv`, `.xlsx`/`.xlsm` | The standard gene presence/absence table is preferred. `gene_presence_absence.Rtab` is a presence/absence matrix and does not contain the same annotation metadata as the CSV. |
| **eggNOG-mapper** | `*.emapper.annotations` | `.tsv`, `.csv`, `.xlsx`/`.xlsm` | Metadata/comment preambles are skipped and the `#query` header is preserved. Column order can vary by eggNOG-mapper version. |
| **KEGG / BlastKOALA** | Downloaded tab-delimited result | `.txt`, `.tsv`, `.csv`, `.xlsx`/`.xlsm` | Supports headerless ID-only lists and the usual 2-column sequence-ID/K-number or 3-column sequence-ID/K-number/definition result. Headered tables with recognized ID/KO/annotation columns are also supported. |
| **antiSMASH** | Region-summary table with region, type, and coordinates | `.tsv`, `.csv`, `.xlsx`/`.xlsm` | Supply a tabular summary with start/end (e.g. From/To) columns. The HTML report and region GenBank files are not parsed directly by this version. |

The parser does **not** directly read every output produced by each tool. In particular, use the structured tables above rather than HTML reports, and use an antiSMASH region-summary table rather than raw GenBank/JSON output. If a source does not expose a suitable tabular file, export/prepare a table with the required identifiers and fields first.

## Requirements

- Python 3.9 or newer recommended
- `pandas`
- `openpyxl` for Excel input

The script checks for `pandas` and `openpyxl` and attempts to install missing packages using the active Python interpreter. This requires a working `pip` and internet access. Manual installation:

```bat
py -3 -m pip install pandas openpyxl
```

## Usage (Windows Command Prompt)

Place `ai-parser.py` and your input files in a folder. Individual source inputs are optional except that either Bakta or Prokka is required to create the master gene list.

```bat
py -3 ai-parser.py --bakta "bakta.tsv" --roary "gene_presence_absence.csv" --eggnog "sample.emapper.annotations" --kegg "blastkoala.txt" --antismash "regions.tsv" --out "merged_gene_annotations.tsv"
```

A Prokka-based run can instead use:

```bat
py -3 ai-parser.py --prokka "prokka.tsv" --eggnog "sample.emapper.annotations" --out "merged_gene_annotations.tsv"
```

For example, to use an Excel export for a source:

```bat
py -3 ai-parser.py --bakta "bakta.xlsx" --roary "roary.xlsx" --eggnog "eggnog.xlsx" --out "merged_gene_annotations.tsv"
```

The original multi-sheet workbook remains supported:

```bat
py -3 ai-parser.py --input "Legacy workbook.xlsx" --out "merged_gene_annotations.tsv"
```

Do not combine `--input` with individual source options. Use quotation marks around paths containing spaces. If `py` is not recognized, try `python` instead.

## Output and matching

The output defaults to `merged_gene_annotations.tsv`. It contains one row per unique Bakta/Prokka locus tag, plus any rows retained from an optional legacy `Output` worksheet. Annotation columns are appended for Roary, eggNOG, KEGG, and antiSMASH.

- **Bakta/Prokka:** supplies the master CDS identifiers and, when available, contig, feature, strand, coordinates, gene, and product.
- **Roary:** joins gene clusters using isolate/GFF cells that contain matching locus tags.
- **eggNOG-mapper:** joins by query/locus tag, including normalization of common variants such as `P1_1_00959` to `P1_00959`.
- **KEGG/BlastKOALA:** imports K numbers and definitions when present; an ID-only list can only report identifier membership.
- **antiSMASH:** associates CDS with region(s) by numeric coordinate overlap. This implementation does not verify contig identity, so only use region coordinates from the same contig/coordinate context as the master CDS list. Coordinate overlap is not proof of biological function.

Multiple matches are combined into a single cell. Missing optional files are skipped. A missing or unrecognized source header produces a warning/error instead of an assumed match.

## Troubleshooting

| Problem | Check |
|---|---|
| Python launcher not recognized | Install Python and enable PATH, or use `python` instead of `py -3`. |
| Dependency installation fails | Check internet access and run the manual pip command above. |
| No gene/CDS rows | Provide a usable Bakta or Prokka feature table. |
| Roary annotations are blank | Check that isolate/GFF cells contain the same locus tags as the master list. |
| eggNOG annotations are blank | Check that the `#query` values match master locus tags after normalization. |
| KEGG has no KO annotations | Confirm that the supplied result includes K numbers/definitions, not only IDs. |
| antiSMASH annotations are blank | Supply a region summary with numeric From/To or Start/End columns and compatible coordinates. |
| Source not identified | Check the error's displayed headers and supply the native table or a supported Excel export. |

## Project files

- `ai-parser.py` — main parser
- `README.md` — project overview and usage
- `ai-parser_readme_supported_layout.md` — detailed format and matching notes

© Z. B. Lara, 2026
