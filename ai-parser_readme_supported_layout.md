# AI-Parser — Supported input formats and matching details

**Documentation for AI-Parser `0.2.0` (development version).** The GitHub release `0.1.2` is a separate, unchanged release until updated and tagged. Run `py -3 ai-parser.py --version` to check the version of this script.

This guide describes the individual-file workflow in `ai-parser.py`. Each annotation source can be supplied separately, and each source is optional except for the master gene list (**Bakta or Prokka**). If both are supplied, Bakta is used as the master list. Excel `.xlsx`/`.xlsm` exports are supported for each source.

## 1. Input options

| Argument | Typical file | Required? | Purpose |
|---|---|---:|---|
| `--bakta` | Bakta `*.tsv` | Alternative master | CDS/features and locus tags; takes precedence over Prokka when both are provided |
| `--prokka` | Prokka `*.tsv` | Alternative master | CDS/features and locus tags |
| `--roary` | `gene_presence_absence.csv` | No | Pangenome clusters and isolate-level gene IDs |
| `--eggnog` | `*.emapper.annotations` | No | Functional annotations keyed by `#query` |
| `--kegg` | BlastKOALA downloaded `.txt`/TSV | No | Sequence IDs and available K numbers/definitions |
| `--antismash` | Region-summary TSV/CSV | No | Biosynthetic gene cluster (BGC) region coordinates and descriptions |
| `--input` | Legacy multi-sheet `.xlsx` workbook | Alternative mode | Reads the earlier workbook layout; cannot be combined with the individual source arguments |
| `--out` | `merged_gene_annotations.tsv` | No | Output path |

Example:

```bat
py -3 ai-parser.py --bakta "bakta.tsv" --roary "gene_presence_absence.csv" --eggnog "sample.emapper.annotations" --kegg "blastkoala.txt" --antismash "regions.tsv" --out "merged_gene_annotations.tsv"
```

Minimal example:

```bat
py -3 ai-parser.py --prokka "prokka.tsv" --out "merged_gene_annotations.tsv"
```

## 2. Formats by tool

### Bakta

Preferred: main Bakta TSV output, typically `<prefix>.tsv`. Headered tables are expected to include a locus-tag field and ideally sequence/contig, feature type, start, stop/end, strand, gene, and product. The GFF3 output is also supported for CDS features. Excel exports are supported.

### Prokka

Preferred: Prokka `<prefix>.tsv`, with common columns such as `locus_tag`, `ftype`, `length_bp`, `gene`, `EC_number`, `COG`, and `product`. GFF/GFF3 is also supported for CDS features. Excel exports are supported. The legacy workbook reader additionally accepts a headerless seven-column layout in the order: `contig, locus_tag, feature, start, end, strand, product`.

### Roary

Preferred: `gene_presence_absence.csv`, with metadata columns such as `Gene`, `Non-unique Gene name`, `Annotation`, `No. isolates`, `No. sequences`, `Avg sequences per isolate`, `Genome Fragment`, and `Order within Fragment`, followed by isolate-specific gene ID cells. IDs in those cells must match Bakta/Prokka locus tags after normalization.

`gene_presence_absence.Rtab` is a binary presence/absence matrix. It is not a complete replacement for the standard CSV because it may not include the same gene metadata or source locus-tag cells required by this merger. Excel exports are supported if the relevant headers and ID cells are preserved.

### eggNOG-mapper

Preferred: `.emapper.annotations` (tab-delimited). These files often start with comment/metadata lines and then a `#query` header. The parser skips metadata before that header and reads the query field. Common annotation fields include `Preferred_name`, `Description`, `COG_category`, `EC`, `eggNOG_OGs`, `KEGG_ko`, and `KEGG_Pathway`; the exact field names/order can vary by version. Excel exports are supported if the query identifier column is retained.

### KEGG / BlastKOALA

Preferred: downloaded tab-delimited BlastKOALA result. The parser supports:

- A one-column, headerless sequence/CDS ID list.
- A two-column, headerless sequence ID and K number result.
- A three-column, headerless sequence ID, K number, and definition result.
- A headered table with recognizable sequence/CDS ID, KO/K number, and annotation/definition columns.
- Excel exports of these tables.

An ID-only list can only indicate whether an ID occurs in the list. It cannot supply K numbers, pathways, or definitions that are absent from the input.

### antiSMASH

Preferred: a region-summary table with a region identifier, biosynthetic type, and numeric start/end coordinates. Recognized coordinate headers include `From`/`To`, `Start`/`End`, and `Region From`/`Region To`. Optional columns include similarity confidence and most similar known cluster. CSV/TSV and Excel summary tables are supported.

This version does **not** directly parse the antiSMASH HTML report, region GenBank files, or JSON output. Prepare or export a tabular region summary before supplying antiSMASH data.

## 3. Header detection and unsupported inputs

The program uses file extensions, source-specific header fields, and content to identify supported inputs. eggNOG comment lines are handled separately so that `#query` remains a usable header. KEGG/BlastKOALA headerless output is supported. If the program cannot confidently identify Roary, eggNOG, or antiSMASH columns, it stops with a message that includes the observed column headers rather than silently interpreting an unrelated file.

Supported tabular extensions include `.tsv`, `.tab`, `.csv`, `.txt`, `.annotations`, and `.out` where appropriate. Excel `.xlsx`/`.xlsm` is supported for all sources. GFF/GFF3 is supported for Bakta/Prokka CDS features. Other formats are not guaranteed to be read directly.

## 4. Matching rules and output

- Bakta or Prokka supplies the preferred master list of locus tags and genomic coordinates.
- IDs are normalized by trimming whitespace and ignoring case.
- A common eggNOG-style `prefix_1_number` pattern is normalized (example: `P1_1_00959` → `P1_00959`).
- Roary is matched through IDs in isolate/GFF columns.
- KEGG/BlastKOALA annotations are keyed by the sequence/CDS ID; K numbers and definitions are imported when available.
- antiSMASH matching uses numeric coordinate overlap. The current implementation does not verify contig identity during overlap checks. CDS and BGC coordinates must therefore share the same contig and coordinate context for the result to be meaningful.
- Multiple annotations are combined into one cell, preserving one row per unique master locus tag.
- The output is a UTF-8, tab-separated `.tsv` file. In Excel, select **Tab** as the delimiter if prompted.

## 5. Requirements and troubleshooting

Requirements: Python 3, `pandas`, and `openpyxl` for Excel input. The script checks for missing dependencies and attempts installation with the active Python interpreter; internet access and `pip` are required for automatic installation.

```bat
py -3 -m pip install pandas openpyxl
```

| Problem | Check |
|---|---|
| No gene/CDS rows | Supply a valid Bakta or Prokka file with locus tags. |
| Roary annotations are blank | Check isolate/GFF cells for IDs matching the master locus tags. |
| eggNOG annotations are blank | Check the `#query` column and ID normalization. |
| KEGG matches but no KO/pathway details | The source may contain IDs only, or lack recognized annotation columns. |
| antiSMASH annotations are blank | Check numeric region coordinates and coordinate/contig compatibility. |
| Source headers not recognized | Use the native tool output described above and retain its column headers. |
| Input path not found | Confirm the path and quote names containing spaces. |

© Z. B. Lara, 2026
