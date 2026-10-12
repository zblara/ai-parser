# AI-Parser — Annotation Intermerger

Merge Roary, Prokka, eggNOG, KEGG, and antiSMASH annotations into a single TSV table with **one row per Prokka CDS**.

## 1. Supported input format

The parser reads **one Excel workbook (`.xlsx`)** containing some or all of the worksheets listed below. Worksheet names are case-insensitive, but should otherwise use the names shown here.

| Worksheet | Required? | Header row? | Purpose |
|---|---:|---|---|
| `Prokka` | Strongly recommended | Either supported format below | Defines the main CDS/gene rows and supplies locus tags and coordinates |
| `Roary` | No | Yes | Supplies gene-cluster and isolate/presence metadata |
| `eggnog` | No | Yes | Supplies functional annotations keyed by locus tag/query |
| `KEGG` | No | Optional / ID-only is supported | Supplies CDS IDs and, if present, additional annotation columns |
| `antiSMASH` | No | Yes | Supplies biosynthetic gene cluster (BGC) region summaries and coordinates |
| `Output` | No | Yes | Optional pre-existing annotation table; its columns and curated values are preserved where possible |

The parser can run with missing optional worksheets, but annotations from a missing sheet cannot be added. `Prokka` is the preferred master list. If no Prokka records are found, the parser falls back to IDs in `Output`.

### 1.1 Prokka worksheet

The parser supports either a **headerless seven-column layout** or a **headered table**.

**Option A — headerless (the sample workbook format):** The columns must be in this exact order, with no header row.

| Column position | Meaning | Example |
|---:|---|---|
| 1 | Contig / sequence ID | `contig00001` |
| 2 | Locus tag / Prokka CDS ID | `P1_00959` |
| 3 | Feature type | `CDS` |
| 4 | Start coordinate | `1200` |
| 5 | End coordinate | `1850` |
| 6 | Strand | `+` or `-` |
| 7 | Product / annotation | `hypothetical protein` |

Example (tab-separated cells in Excel):

```text
contig00001    P1_00959    CDS    1200    1850    +    hypothetical protein
contig00001    P1_00960    CDS    1900    2400    -    ATP-binding protein
```

Do not add a title row above the data. The parser treats the first row as data when it does not detect a recognized header.

**Option B — headered:** Include a header row with recognizable column names. Recommended headers:

```text
seqid    locus_tag    ftype    start    end    strand    product
```

Recognized alternatives include `Prokka Code` or `gene_id` for the locus tag; `contig`, `sequence`, or `accession` for the contig; `feature` or `type` for feature type; `begin` / `stop` for coordinates; and `annotation` / `description` for product. A locus-tag column is required.

### 1.2 Roary worksheet

Use a Roary `gene_presence_absence`-style table with a header row.

The parser reads the following metadata columns when present:

| Recognized column | Output field |
|---|---|
| `Gene` | Roary gene cluster ID |
| `Non-unique Gene name` | Non-unique gene name |
| `Annotation` | Cluster annotation |
| `No. isolates` | Number of isolates |
| `No. sequences` | Number of sequences |
| `Avg sequences per isolate` | Average sequences per isolate |
| `Genome Fragment` | Genome fragment |
| `Order within Fragment` | Order within fragment |

For per-CDS matching, the parser searches columns whose names contain `GFF` or start with `Prokka_`. Standard Roary tables without such column names use columns after the first 14 metadata columns as isolate/GFF columns. Cells may contain semicolon-separated IDs.

**Important:** The IDs in the Roary GFF/isolate columns need to correspond to Prokka locus tags (or their normalized equivalents) for the join to work.

### 1.3 eggNOG worksheet

Use an eggNOG-mapper annotation table with a header row. The most important field is the locus tag/query identifier.

| Information | Recognized headers (examples) |
|---|---|
| Query / locus tag (required for matching) | `Locus_Tag`, `locus_tag`, `query`, `query_name`, `#query`, `protein_id` |
| Product | `Product`, `product` |
| Gene name | `Gene`, `Preferred_name`, `preferred_name` |
| Enzyme Commission number | `EC_number`, `EC number`, `EC` |
| Database / annotation level | `Eggnog_Db`, `EggNOG DB`, `max_annot_lvl` |
| Source / ortholog group | `Source_ID`, `COG`, `eggNOG_OGs` |
| Description | `Description`, `description` |
| Functional category | `Eggnog_category`, `COG_category`, `category` |
| Identity | `Identity`, `identity` |
| E-value | `e-value`, `evalue`, `E-value` |
| Bit score | `bitscore`, `bit score`, `score` |
| GC content (if supplied) | `G+C%`, `GC`, `GC%` |

Column names are matched flexibly by ignoring case, spaces, and punctuation. The query/locus-tag column must be present; otherwise eggNOG matching is skipped with a warning.

A common tag variant is normalized for matching: `P1_1_00959` can match `P1_00959`. Other tag formats are not automatically guaranteed to match.

### 1.4 KEGG worksheet

Two formats are supported:

- **ID-only list:** Put one Prokka/CDS ID per row in the first column. A header is not required.
- **Table with annotations:** Put the CDS ID in the first column and any available KEGG annotation fields in later columns. The parser reports whether additional annotation columns contain data, but it does not currently map individual KO, pathway, or enzyme fields into dedicated output columns.

Example ID-only list:

```text
P1_00959
P1_00960
P1_00961
```

The parser reports whether each Prokka ID is present in this list. If the sheet contains only IDs, the output will explicitly say that KO/pathway/annotation columns were not supplied; it does not infer them.

### 1.5 antiSMASH worksheet

Use a region-summary table with a header row and **region start/end coordinates**. Recognized headers include:

| Information | Recognized headers (examples) |
|---|---|
| Region ID | `Region`, `Region ID` |
| Biosynthetic type | `Type`, `Product`, `Biosynthetic type` |
| Region start | `From`, `Start`, `Region From` |
| Region end | `To`, `End`, `Region To` |
| Similarity confidence | `Similarity Confidence`, `Confidence` |
| Closest known cluster | `Most similar known cluster`, `Known cluster`, `Best hit` |
| Subtype/class (optional) | `Subtype`, `Class`, `Additional type` |

The parser associates a CDS with a region when their numeric coordinates overlap. This is a **coordinate-overlap match**, not proof that a gene is functionally part of a biosynthetic pathway. Coordinates must be in the same coordinate system and sequence/contig context for the match to be meaningful. The current implementation checks coordinate overlap and does not verify contig names.

### 1.6 Optional Output worksheet

The optional `Output` sheet should have a header row and a column identifying the gene, preferably `Prokka Code`. Recognized ID alternatives include `locus_tag` and `gene_id`.

Existing values are retained where possible; missing values may be filled from parsed sources. Prokka defines the main row set when it contains records. Existing `Output` rows that do not occur in Prokka are retained at the end.

## 2. Matching rules and output

- The main row set is based on unique Prokka locus tags when available.
- IDs are normalized by trimming whitespace and ignoring case.
- The specific eggNOG-style `prefix_1_number` pattern is normalized (for example, `P1_1_00959` → `P1_00959`).
- Roary is matched through IDs in its GFF/isolate columns.
- KEGG matching checks membership in the provided ID list.
- antiSMASH matching checks coordinate overlap with each listed BGC region.
- Multiple matches are combined into a single cell using semicolons or separators, so the output remains one row per Prokka ID.
- The parser writes a UTF-8, tab-separated `.tsv` file. Open it in Excel using the **Tab** delimiter if prompted.

## 3. Windows installation and use

### Requirements

- Python 3 installed on Windows.
- The script needs `pandas` and `openpyxl`. The Windows-enabled version can check for these and attempt to install missing packages, provided `pip` and internet access are available.

### Run from Command Prompt

Place the Python script and workbook in the same folder, for example:

```text
C:\AI-Parser\
    ai-parser-windows.py
    Sample data.xlsx
```

Open Command Prompt and run:

```bat
cd /d C:\AI-Parser
py -3 ai-parser-windows.py --input "Sample data.xlsx" --out merged_gene_annotations.tsv
```

If the Windows Python launcher (`py`) is not available, try:

```bat
python ai-parser-windows.py --input "Sample data.xlsx" --out merged_gene_annotations.tsv
```

### Optional batch launcher

Keep `run-ai-parser.bat` in the same folder as `ai-parser-windows.py`. Then run:

```bat
run-ai-parser.bat "Sample data.xlsx"
```

To set an output filename:

```bat
run-ai-parser.bat "Sample data.xlsx" "final_annotations.tsv"
```

Use quotation marks around file or folder paths that contain spaces.

## 4. Troubleshooting

| Message / issue | What to check |
|---|---|
| `'py' is not recognized` | Install Python and enable PATH, or try `python` instead. |
| `No module named pip` | Repair the Python installation or enable/install `pip`. |
| Package installation fails | Check internet access, proxy settings, and package permissions. |
| Input workbook does not exist | Confirm the path and filename; include quotation marks around names containing spaces. |
| Roary annotations are blank | Confirm the IDs in the Roary GFF/isolate columns match the Prokka locus tags. |
| eggNOG annotations are blank | Confirm a recognized query/locus-tag column exists and the IDs match after normalization. |
| antiSMASH annotations are blank | Confirm the region summary includes numeric start/end coordinates and that CDS and region coordinates use the same coordinate system. |
| KEGG shows ID matches but no pathway data | The KEGG sheet may contain only IDs. Add annotation columns to the source if available; the parser does not infer missing KO/pathway annotations. |
| Unexpected columns or missing annotations | Confirm worksheet names and headers follow the supported layouts above. |

## 5. Example command

```bat
py -3 ai-parser-windows.py --input "Sample data.xlsx" --out merged_gene_annotations.tsv
```

The program prints the output path, number of gene/CDS rows, and number of columns when it finishes successfully.

---

©2026 Zachary B. Lara
