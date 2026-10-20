#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple
import importlib.util
import subprocess


__version__ = "0.2.0"


def ensure_dependencies() -> None:
    """Install required third-party packages into this Python environment if missing."""
    required = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
    }
    missing = [package for module, package in required.items()
               if importlib.util.find_spec(module) is None]
    if not missing:
        print("Dependency check: pandas and openpyxl are already installed.")
        return

    print("Dependency check: missing package(s): " + ", ".join(missing))
    print("Attempting to install them with this Python interpreter...\n")
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "--user", *missing
        ])
    except subprocess.CalledProcessError as exc:
        print("\nERROR: Automatic dependency installation failed.", file=sys.stderr)
        print("Check your internet connection and that pip is available.", file=sys.stderr)
        print("You can try manually running:", file=sys.stderr)
        print(f'  "{sys.executable}" -m pip install --user ' + " ".join(missing), file=sys.stderr)
        raise SystemExit(exc.returncode)


ensure_dependencies()
import pandas as pd


print(f'''
********************************************************
𝓐𝓷𝓷𝓸𝓽𝓪𝓽𝓲𝓸𝓷 𝓘𝓷𝓽𝓮𝓻𝓶𝓮𝓻𝓰𝓮𝓻 - 𝓟𝓪𝓻𝓼𝓮𝓻 𝓟𝓻𝓸𝓰𝓻𝓪𝓶 (𝒜𝐼-𝒫𝒶𝓇𝓈𝑒𝓇) — v{__version__}
********************************************************

Merge Roary, Prokka, eggNOG, KEGG and antiSMASH annotations into one TSV row per CDS.

Accepts separate optional files from Bakta or Prokka, Roary, eggNOG-mapper,
KEGG/BlastKOALA, and antiSMASH. Excel workbooks are also accepted for each source.
Use Bakta or Prokka as the master CDS list; all other sources are optional.

Examples (Windows Command Prompt):
    py ai-parser.py --prokka prokka.tsv --roary gene_presence_absence.csv \
      --eggnog sample.emapper.annotations --kegg blastkoala.txt \
      --antismash regions.tsv --out merged_gene_annotations.tsv
    py ai-parser.py --bakta bakta.tsv --eggnog annotations.xlsx
    py ai-parser.py --input "Legacy workbook.xlsx"

Dependencies: pandas and openpyxl (automatically checked/installed when needed).
''')


BASE_OUTPUT_COLUMNS = [
    "Prokka Code",
    "KEGG No",
    "Gene Name",
    "KEGG Annotation",
    "Start",
    "End",
    "G+C%",
    "COG",
    "COG Functional Classification",
    "Associated with which BGC (based on antiSMASH result)",
    "Description Prokka",
]

DERIVED_COLUMNS = [
    "Roary Gene",
    "Roary Non-unique Gene name",
    "Roary Annotation",
    "Roary No. isolates",
    "Roary No. sequences",
    "Roary Avg sequences per isolate",
    "Roary Genome Fragment",
    "Roary Order within Fragment",
    "Roary Prokka GFF column matched",
    "Prokka Contig",
    "Prokka Feature",
    "Prokka Strand",
    "Prokka Product",
    "eggNOG Locus_Tag",
    "eggNOG Product",
    "eggNOG Gene",
    "eggNOG EC_number",
    "eggNOG DB",
    "eggNOG Source_ID",
    "eggNOG Description",
    "eggNOG Category",
    "eggNOG Identity",
    "eggNOG e-value",
    "eggNOG bitscore",
    "KEGG ID list match",
    "KEGG annotation availability",
    "antiSMASH Region",
    "antiSMASH Type",
    "antiSMASH Region From",
    "antiSMASH Region To",
    "antiSMASH Similarity Confidence",
    "antiSMASH Most similar known cluster",
    "Match basis",
    "CDS overlaps antiSMASH region",
]


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


def text(value: Any) -> str:
    """Convert a cell to clean text, avoiding strings like '123.0' for integers."""
    if is_blank(value):
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    result = str(value).replace("\xa0", " ").strip()
    return result


def header_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", text(value).casefold())


def normalise_locus_tag(value: Any) -> str:
    """Normalise common eggNOG tag variants such as P1_1_00959 -> P1_00959."""
    raw = text(value)
    if not raw:
        return ""
    raw = re.sub(r"\s+", "", raw)
    # The sample eggNOG output inserts '_1_' between the prefix and numeric tag.
    raw = re.sub(r"^(.+?)_1_(\d+)$", r"\1_\2", raw, flags=re.IGNORECASE)
    return raw.casefold()


def find_column(columns: Iterable[Any], aliases: Sequence[str]) -> Optional[str]:
    """Find a column by normalized name, checking aliases in priority order."""
    columns = [text(c) for c in columns]
    lookup = {header_key(c): c for c in columns}
    for alias in aliases:
        key = header_key(alias)
        if key in lookup:
            return lookup[key]
    return None


def first_nonblank(row: Mapping[str, Any], aliases: Sequence[str], default: Any = "") -> Any:
    col = find_column(row.keys(), aliases)
    if col is not None and not is_blank(row.get(col)):
        return row.get(col)
    return default


def split_gene_ids(value: Any) -> List[str]:
    """Roary cells may contain one or more semicolon-separated locus tags."""
    raw = text(value)
    if not raw or raw in {"-", "NA", "nan"}:
        return []
    ids = []
    for item in re.split(r"[;\n\r]+", raw):
        item = item.strip()
        if item and item not in {"-", "NA", "nan"}:
            ids.append(item)
    return ids


def cell_number(value: Any) -> Optional[int]:
    raw = text(value).replace(",", "")
    if not raw:
        return None
    try:
        return int(float(raw))
    except (TypeError, ValueError):
        return None


def unique_values(values: Iterable[Any]) -> List[Any]:
    out: List[Any] = []
    seen = set()
    for value in values:
        if is_blank(value):
            continue
        rendered = text(value)
        if rendered not in seen:
            seen.add(rendered)
            out.append(value)
    return out


def merge_values(values: Iterable[Any]) -> Any:
    vals = unique_values(values)
    if not vals:
        return ""
    if len(vals) == 1:
        return vals[0]
    return "; ".join(text(v) for v in vals)


def _read_delimited(path: Path, kind: str) -> pd.DataFrame:
    """Read CSV/TSV-like files, handling eggNOG's comment preamble and header."""
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    lines = raw.splitlines()
    if kind == "eggnog":
        # eggNOG-mapper files contain metadata comments; preserve the #query header.
        header_index = next((i for i, line in enumerate(lines)
                             if line.lstrip().casefold().startswith("#query\t") or
                                line.lstrip().casefold().startswith("#query,")), None)
        if header_index is not None:
            lines = lines[header_index:]
        else:
            lines = [line for line in lines if line.strip() and not line.lstrip().startswith("##")]
    else:
        lines = [line for line in lines if line.strip() and not line.lstrip().startswith("##")]
    if not lines:
        return pd.DataFrame()
    sample = "\n".join(lines[:8])
    sep = "," if path.suffix.casefold() == ".csv" else "\t"
    if path.suffix.casefold() not in {".csv", ".tsv", ".tab"}:
        # .emapper.annotations and .txt are commonly tab-separated; infer if needed.
        sep = "\t" if sample.count("\t") >= sample.count(",") else ","
    from io import StringIO
    header = 0
    if kind == "kegg":
        first_cells = [header_key(v) for v in next(csv.reader([lines[0]], delimiter=sep))]
        header_tokens = {"id", "geneid", "locustag", "prokkacode", "query", "keggid", "gene", "cdsid", "keggno", "keggannotation", "sequenceid"}
        if not (set(first_cells) & header_tokens):
            header = None  # KEGG/BlastKOALA downloads and ID lists are often headerless.
    return pd.read_csv(StringIO("\n".join(lines)), sep=sep, header=header, dtype=object,
                       keep_default_na=False, engine="python", on_bad_lines="warn")


def read_source_file(path: Optional[Path], kind: str) -> Optional[pd.DataFrame]:
    """Read a source file independently; infer CSV/TSV/XLSX and common text formats."""
    if path is None:
        return None
    if not path.is_file():
        raise FileNotFoundError(f"{kind} input does not exist: {path}")
    suffix = path.suffix.casefold()
    if suffix in {".xlsx", ".xlsm"}:
        xls = pd.ExcelFile(path, engine="openpyxl")
        # Prefer a sheet whose name identifies the source, otherwise the first sheet.
        candidates = {name.casefold(): name for name in xls.sheet_names}
        wanted = {"eggnog": ["eggnog", "annotations"], "antismash": ["antismash", "regions"],
                  "prokka": ["prokka", "features"], "bakta": ["bakta", "features"],
                  "roary": ["roary", "gene_presence_absence"], "kegg": ["kegg", "blastkoala"]}.get(kind, [])
        sheet = next((candidates[n] for n in wanted if n in candidates), xls.sheet_names[0])
        if kind in {"prokka", "bakta"}:
            # Headered tool output is normal; preserve headerless legacy 7-column extracts.
            df0 = pd.read_excel(xls, sheet_name=sheet, header=0, dtype=object)
            cols = {header_key(c) for c in df0.columns}
            if cols & {"locustag", "prokkacode", "seqid", "sequenceid", "start"}:
                df = df0
            else:
                df = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
        elif kind == "kegg":
            df = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
        else:
            df = pd.read_excel(xls, sheet_name=sheet, header=0, dtype=object)
    elif suffix in {".csv", ".tsv", ".tab", ".txt", ".annotations", ".out"}:
        df = _read_delimited(path, kind)
    elif suffix in {".gff", ".gff3"} and kind in {"prokka", "bakta"}:
        rows = []
        for line in path.read_text(encoding="utf-8-sig", errors="replace").splitlines():
            if not line or line.startswith("#"):
                continue
            parts = line.split("\t")
            if len(parts) != 9 or parts[2].casefold() != "cds":
                continue
            attrs = dict((kv.split("=", 1) + [""])[:2] for kv in parts[8].split(";") if "=" in kv)
            rows.append({"contig": parts[0], "feature": parts[2], "start": parts[3], "end": parts[4],
                         "strand": parts[6], "locus_tag": attrs.get("locus_tag", attrs.get("ID", "")),
                         "gene": attrs.get("gene", ""), "product": attrs.get("product", "")})
        df = pd.DataFrame(rows)
    else:
        raise ValueError(f"Unsupported {kind} file format '{suffix}'. Use XLSX, CSV/TSV/TXT; Prokka/Bakta also support GFF/GFF3.")
    df = df.dropna(how="all").reset_index(drop=True)
    if kind in {"prokka", "bakta"}:
        df.columns = [text(c) for c in df.columns]
    return df


def read_workbook_sheets(path: Path) -> Dict[str, pd.DataFrame]:
    """Backward-compatible reader for the original multi-sheet workbook."""
    xls = pd.ExcelFile(path, engine="openpyxl")
    available = {name.casefold(): name for name in xls.sheet_names}
    sheets: Dict[str, pd.DataFrame] = {}
    for kind in ("output", "roary", "eggnog", "antismash"):
        real = available.get(kind)
        if real:
            df = pd.read_excel(xls, sheet_name=real, header=0, dtype=object)
            df.columns = [text(c) for c in df.columns]
            sheets[kind] = df.dropna(how="all").reset_index(drop=True)
    for kind in ("prokka", "kegg"):
        real = available.get(kind)
        if real:
            df = pd.read_excel(xls, sheet_name=real, header=None, dtype=object)
            sheets[kind] = df.dropna(how="all").reset_index(drop=True)
    return sheets


def parse_prokka(df: Optional[pd.DataFrame]) -> Dict[str, Dict[str, Any]]:
    """Return records keyed by normalized locus tag; supports headerless sample format."""
    records: Dict[str, Dict[str, Any]] = {}
    if df is None or df.empty:
        return records

    first_row = [text(v) for v in df.iloc[0].tolist()]
    first_keys = {header_key(v) for v in first_row}
    column_keys = {header_key(v) for v in df.columns}
    header_tokens = {"locustag", "prokkacode", "ftype", "seqid", "sequenceid", "start"}
    direct_headered = bool(column_keys & header_tokens) and not all(str(c).isdigit() for c in df.columns)
    headered = direct_headered or bool(first_keys & header_tokens)

    if headered:
        if direct_headered:
            headers = [text(v) for v in df.columns]
            data = df.copy()
        else:
            headers = first_row
            data = df.iloc[1:].copy()
            data.columns = headers
        locus_col = find_column(data.columns, ["locus_tag", "Locus Tag", "Prokka Code", "gene_id", "ID", "locus_tag_id"])
        contig_col = find_column(data.columns, ["contig", "seqid", "sequence", "Sequence Id", "accession"])
        feature_col = find_column(data.columns, ["ftype", "feature", "type"])
        start_col = find_column(data.columns, ["start", "begin"])
        end_col = find_column(data.columns, ["end", "stop"])
        strand_col = find_column(data.columns, ["strand"])
        product_col = find_column(data.columns, ["product", "annotation", "description", "Product"])
        gene_col = find_column(data.columns, ["gene"])
        if locus_col is None:
            raise ValueError("The Prokka sheet appears to have headers, but no locus_tag/Prokka Code column was found.")

        for _, row in data.iterrows():
            locus = text(row.get(locus_col))
            key = normalise_locus_tag(locus)
            if not key:
                continue
            records[key] = {
                "locus_tag": locus,
                "contig": text(row.get(contig_col)) if contig_col else "",
                "feature": text(row.get(feature_col)) if feature_col else "CDS",
                "start": cell_number(row.get(start_col)) if start_col else None,
                "end": cell_number(row.get(end_col)) if end_col else None,
                "strand": text(row.get(strand_col)) if strand_col else "",
                "product": text(row.get(product_col)) if product_col else "",
                "gene": text(row.get(gene_col)) if gene_col else "",
            }
    else:
        # Sample schema: contig, locus_tag, feature, start, end, strand, product.
        for _, row in df.iterrows():
            vals = row.tolist()
            if len(vals) < 2:
                continue
            locus = text(vals[1])
            key = normalise_locus_tag(locus)
            if not key:
                continue
            records[key] = {
                "locus_tag": locus,
                "contig": text(vals[0]) if len(vals) > 0 else "",
                "feature": text(vals[2]) if len(vals) > 2 else "CDS",
                "start": cell_number(vals[3]) if len(vals) > 3 else None,
                "end": cell_number(vals[4]) if len(vals) > 4 else None,
                "strand": text(vals[5]) if len(vals) > 5 else "",
                "product": text(vals[6]) if len(vals) > 6 else "",
                "gene": "",
            }
    return records


def parse_roary(df: Optional[pd.DataFrame]) -> Dict[str, List[Dict[str, Any]]]:
    """Map every Prokka locus tag in Roary's isolate/GFF columns to its cluster row."""
    result: Dict[str, List[Dict[str, Any]]] = {}
    if df is None or df.empty:
        return result

    gene_col = find_column(df.columns, ["Gene"])
    gff_columns = [c for c in df.columns if "gff" in text(c).casefold() or text(c).casefold().startswith("prokka_")]
    if not gff_columns:
        # Standard Roary's first 14 columns are cluster metadata; remaining columns are isolate IDs.
        gff_columns = list(df.columns[14:])

    metadata_aliases = {
        "gene": ["Gene"],
        "non_unique_gene_name": ["Non-unique Gene name"],
        "annotation": ["Annotation"],
        "no_isolates": ["No. isolates"],
        "no_sequences": ["No. sequences"],
        "avg_sequences_per_isolate": ["Avg sequences per isolate"],
        "genome_fragment": ["Genome Fragment"],
        "order_within_fragment": ["Order within Fragment"],
    }

    for _, row in df.iterrows():
        row_dict = row.to_dict()
        info: Dict[str, Any] = {}
        for output_key, aliases in metadata_aliases.items():
            info[output_key] = first_nonblank(row_dict, aliases, "")
        matched_any = False
        for col in gff_columns:
            for gene_id in split_gene_ids(row.get(col)):
                key = normalise_locus_tag(gene_id)
                if not key:
                    continue
                entry = dict(info)
                entry["matched_gff_column"] = text(col)
                entry["matched_gene_id"] = gene_id
                result.setdefault(key, []).append(entry)
                matched_any = True
        # Roary clusters with no gene IDs are intentionally ignored for per-CDS merging.
    return result


def parse_eggnog(df: Optional[pd.DataFrame]) -> Dict[str, List[Dict[str, Any]]]:
    """Map eggNOG-style annotations by locus tag/query; normalizes the sample's P1_1_00959 IDs."""
    result: Dict[str, List[Dict[str, Any]]] = {}
    if df is None or df.empty:
        return result
    locus_col = find_column(df.columns, ["Locus_Tag", "locus_tag", "query", "query_name", "#query", "protein_id"])
    if locus_col is None:
        # Do not mistakenly use the gene-symbol column as an ID; fail clearly instead.
        print("WARNING: eggNOG sheet has no recognized Locus_Tag/query column; skipping eggNOG joins.", file=sys.stderr)
        return result

    mapping = {
        "locus_tag": ["Locus_Tag", "locus_tag", "query", "query_name", "#query"],
        "product": ["Product", "product"],
        "gene": ["Gene", "Preferred_name", "preferred_name"],
        "ec_number": ["EC_number", "EC number", "EC"],
        "db": ["Eggnog_Db", "Eggnog_DB", "EggNOG DB", "max_annot_lvl"],
        "source_id": ["Source_ID", "COG", "eggNOG_OGs"],
        "description": ["Description", "description"],
        "category": ["Eggnog_category", "COG_category", "category"],
        "identity": ["Identity", "identity"],
        "evalue": ["e-value", "evalue", "E-value"],
        "bitscore": ["bitscore", "bit score", "score"],
        "gc": ["G+C%", "GC", "GC%"],
    }

    for _, row in df.iterrows():
        row_dict = row.to_dict()
        raw_locus = text(row.get(locus_col))
        key = normalise_locus_tag(raw_locus)
        if not key:
            continue
        entry: Dict[str, Any] = {"locus_tag": raw_locus}
        for output_key, aliases in mapping.items():
            if output_key == "locus_tag":
                continue
            entry[output_key] = first_nonblank(row_dict, aliases, "")
        result.setdefault(key, []).append(entry)
    return result


def parse_kegg(df: Optional[pd.DataFrame]) -> Tuple[set, bool, Dict[str, Dict[str, Any]]]:
    """Parse KEGG IDs and optional per-CDS fields such as KEGG No and KEGG Annotation.

    Supports either an ID-only/headerless first-column list or a table whose first
    row contains headers. KEGG annotations are keyed by the Prokka/CDS ID.
    """
    ids = set()
    records: Dict[str, Dict[str, Any]] = {}
    if df is None or df.empty:
        return ids, False, records

    data = df.copy()
    first_row = [text(v) for v in data.iloc[0].tolist()]
    normalized_headers = {header_key(v) for v in first_row}
    known_header_tokens = {
        "id", "geneid", "locustag", "prokkacode", "query", "keggid",
        "gene", "cdsid", "locus", "keggno", "keggannotation",
        "annotation", "ko", "koid", "pathway", "pathwayid",
    }
    has_header = bool(normalized_headers & known_header_tokens)
    if has_header:
        headers = []
        for i, value in enumerate(first_row):
            headers.append(value if value else f"column_{i + 1}")
        data = data.iloc[1:].copy()
        data.columns = headers
    else:
        # BlastKOALA downloads are headerless: sequence ID, K number, optional definition.
        # A one-column file is treated as an ID list; 2/3 columns carry KO and definition.
        if len(data.columns) >= 3:
            data.columns = ["Sequence ID", "KEGG No", "KEGG Annotation"] + [f"column_{i + 1}" for i in range(3, len(data.columns))]
        elif len(data.columns) == 2:
            data.columns = ["Sequence ID", "KEGG No"]
        else:
            data.columns = ["Sequence ID"] if len(data.columns) else []

    id_col = find_column(data.columns, [
        "Prokka Code", "locus_tag", "gene_id", "CDS ID", "query",
        "query_name", "#query", "ID", "Gene ID", "KEGG ID", "Sequence ID",
    ])
    # For ID-only lists or unrecognized ID headers, use the first column.
    if id_col is None:
        id_col = data.columns[0] if len(data.columns) else None
    if id_col is None:
        return ids, False, records

    kegg_no_col = find_column(data.columns, [
        "KEGG No", "KEGG Number", "KO", "KO ID", "KEGG Orthology",
    ])
    annotation_col = find_column(data.columns, [
        "KEGG Annotation", "KEGG Description", "KEGG Function",
        "Annotation", "Description", "Definition", "Product",
    ])

    annotation_columns = [
        c for c in data.columns
        if c != id_col and any(not is_blank(v) for v in data[c].tolist())
    ]
    annotations_available = bool(annotation_columns)

    for _, row in data.iterrows():
        raw_id = text(row.get(id_col))
        if not raw_id or header_key(raw_id) in {
            "id", "geneid", "locustag", "prokkacode", "query", "keggid",
            "gene", "cdsid",
        }:
            continue
        key = normalise_locus_tag(raw_id)
        if not key:
            continue
        ids.add(key)
        record = {
            "kegg_no": text(row.get(kegg_no_col)) if kegg_no_col else "",
            "kegg_annotation": text(row.get(annotation_col)) if annotation_col else "",
        }
        if any(record.values()):
            existing = records.setdefault(key, {})
            for field, value in record.items():
                if value and not existing.get(field):
                    existing[field] = value
                elif value and existing.get(field) and value != existing[field]:
                    existing[field] = merge_values([existing[field], value])
    return ids, annotations_available, records


def parse_antismash(df: Optional[pd.DataFrame]) -> List[Dict[str, Any]]:
    """Read a tabular antiSMASH region summary; coordinate columns are expected in contig coordinates."""
    regions: List[Dict[str, Any]] = []
    if df is None or df.empty:
        return regions

    aliases = {
        "region": ["Region", "Region ID"],
        "type": ["Type", "Product", "Biosynthetic type"],
        "start": ["From", "Start", "Region From"],
        "end": ["To", "End", "Region To"],
        "confidence": ["Similarity Confidence", "Confidence"],
        "known_cluster": ["Most similar known cluster", "Known cluster", "Best hit"],
        "subtype": ["Subtype", "Class", "Additional type"],
    }
    for _, row in df.iterrows():
        row_dict = row.to_dict()
        info = {key: first_nonblank(row_dict, names, "") for key, names in aliases.items()}
        info["start_num"] = cell_number(info["start"])
        info["end_num"] = cell_number(info["end"])
        # Ignore totally empty rows or rows without coordinates.
        if not info["region"] and info["start_num"] is None and info["end_num"] is None:
            continue
        if info["start_num"] is None or info["end_num"] is None:
            continue
        info["region"] = re.sub(r"\s+", " ", text(info["region"]))
        regions.append(info)
    return regions


def build_output(
    sheets: Dict[str, pd.DataFrame],
    output_path: Path,
) -> pd.DataFrame:
    prokka_records = parse_prokka(sheets.get("prokka"))
    roary_map = parse_roary(sheets.get("roary"))
    eggnog_map = parse_eggnog(sheets.get("eggnog"))
    kegg_ids, kegg_annotations_available, kegg_map = parse_kegg(sheets.get("kegg"))
    regions = parse_antismash(sheets.get("antismash"))

    # Existing Output columns are preserved when available; Prokka defines the CDS row set.
    existing = sheets.get("output")
    if existing is not None and not existing.empty:
        base_columns = [text(c) for c in existing.columns]
        code_col = find_column(base_columns, ["Prokka Code", "locus_tag", "gene_id"])
        existing_by_id: Dict[str, Dict[str, Any]] = {}
        if code_col:
            for _, row in existing.iterrows():
                code = text(row.get(code_col))
                key = normalise_locus_tag(code)
                if key and key not in existing_by_id:
                    existing_by_id[key] = {text(k): v for k, v in row.to_dict().items()}
    else:
        base_columns = BASE_OUTPUT_COLUMNS.copy()
        existing_by_id = {}

    # Use all Prokka CDS as the master. If no Prokka rows were supplied, fall back to Output.
    master_keys = list(prokka_records.keys())
    if not master_keys:
        master_keys = list(existing_by_id.keys())
    # Preserve any Output-only rows at the end rather than silently dropping them.
    for key in existing_by_id:
        if key not in master_keys:
            master_keys.append(key)

    output_rows: List[Dict[str, Any]] = []
    for key in master_keys:
        prokka = prokka_records.get(key, {})
        locus_tag = text(prokka.get("locus_tag")) or next(
            (text(row.get("Prokka Code")) for row in [existing_by_id.get(key, {})] if text(row.get("Prokka Code"))), ""
        )
        base = dict(existing_by_id.get(key, {}))
        if not locus_tag:
            locus_tag = next((text(v) for k, v in base.items() if header_key(k) in {"prokkacode", "locustag", "geneid"}), "")

        # Ensure baseline row has a consistent Prokka Code field.
        base_code_col = find_column(base.keys(), ["Prokka Code", "locus_tag", "gene_id"])
        if base_code_col:
            base[base_code_col] = locus_tag or base.get(base_code_col, "")
        else:
            base["Prokka Code"] = locus_tag
        if not base_columns:
            base_columns = BASE_OUTPUT_COLUMNS.copy()
        for col in base_columns:
            base.setdefault(col, "")

        # Fill missing baseline coordinates and gene name from Prokka/eggNOG, without overwriting curated values.
        for target, value in (("Start", prokka.get("start")), ("End", prokka.get("end"))):
            target_col = find_column(base.keys(), [target])
            if target_col and is_blank(base.get(target_col)) and value is not None:
                base[target_col] = value
        egg_matches = eggnog_map.get(key, [])
        if not egg_matches:
            egg_matches = []
        gene_name_col = find_column(base.keys(), ["Gene Name", "Gene", "gene_name"])
        if gene_name_col and is_blank(base.get(gene_name_col)):
            gene_val = merge_values([entry.get("gene", "") for entry in egg_matches])
            base[gene_name_col] = gene_val or prokka.get("gene", "")
        gc_col = find_column(base.keys(), ["G+C%", "GC%", "GC"])
        if gc_col and is_blank(base.get(gc_col)):
            base[gc_col] = merge_values([entry.get("gc", "") for entry in egg_matches])

        row_out: Dict[str, Any] = dict(base)
        row_out["Prokka Code"] = locus_tag or text(base.get("Prokka Code", ""))

        # Roary cluster and matched isolate/GFF column.
        roary_matches = roary_map.get(key, [])
        roary_fields = [
            ("Roary Gene", "gene"),
            ("Roary Non-unique Gene name", "non_unique_gene_name"),
            ("Roary Annotation", "annotation"),
            ("Roary No. isolates", "no_isolates"),
            ("Roary No. sequences", "no_sequences"),
            ("Roary Avg sequences per isolate", "avg_sequences_per_isolate"),
            ("Roary Genome Fragment", "genome_fragment"),
            ("Roary Order within Fragment", "order_within_fragment"),
        ]
        for output_col, map_key in roary_fields:
            row_out[output_col] = merge_values([m.get(map_key, "") for m in roary_matches])
        row_out["Roary Prokka GFF column matched"] = merge_values([m.get("matched_gff_column", "") for m in roary_matches])

        # Prokka CDS details.
        row_out["Prokka Contig"] = prokka.get("contig", "")
        row_out["Prokka Feature"] = prokka.get("feature", "")
        row_out["Prokka Strand"] = prokka.get("strand", "")
        row_out["Prokka Product"] = prokka.get("product", "")

        # eggNOG fields.
        egg_fields = [
            ("eggNOG Locus_Tag", "locus_tag"),
            ("eggNOG Product", "product"),
            ("eggNOG Gene", "gene"),
            ("eggNOG EC_number", "ec_number"),
            ("eggNOG DB", "db"),
            ("eggNOG Source_ID", "source_id"),
            ("eggNOG Description", "description"),
            ("eggNOG Category", "category"),
            ("eggNOG Identity", "identity"),
            ("eggNOG e-value", "evalue"),
            ("eggNOG bitscore", "bitscore"),
        ]
        for output_col, map_key in egg_fields:
            row_out[output_col] = merge_values([m.get(map_key, "") for m in egg_matches])

        has_kegg = key in kegg_ids
        kegg_record = kegg_map.get(key, {})
        # Fill KEGG fields from the KEGG worksheet only when the existing Output
        # value is blank, preserving any curated values already present.
        for target, source_key in (
            ("KEGG No", "kegg_no"),
            ("KEGG Annotation", "kegg_annotation"),
        ):
            target_col = find_column(row_out.keys(), [target])
            if target_col is None:
                row_out[target] = kegg_record.get(source_key, "")
            elif is_blank(row_out.get(target_col)) and kegg_record.get(source_key):
                row_out[target_col] = kegg_record[source_key]

        row_out["KEGG ID list match"] = "yes" if has_kegg else "no"
        if kegg_record.get("kegg_no") or kegg_record.get("kegg_annotation"):
            row_out["KEGG annotation availability"] = "KEGG No/Annotation imported from KEGG sheet"
        else:
            row_out["KEGG annotation availability"] = (
                "Other annotation columns supplied; no recognized KEGG No/Annotation value for this CDS"
                if kegg_annotations_available
                else "IDs only; no KO/pathway/annotation columns in supplied KEGG sheet"
            )

        # Match CDS coordinates to any antiSMASH regions that overlap them.
        start = prokka.get("start")
        end = prokka.get("end")
        if start is None or end is None:
            start = cell_number(first_nonblank(row_out, ["Start"], ""))
            end = cell_number(first_nonblank(row_out, ["End"], ""))
        overlapping = []
        if start is not None and end is not None:
            cds_start, cds_end = min(start, end), max(start, end)
            for region in regions:
                if cds_start <= region["end_num"] and cds_end >= region["start_num"]:
                    overlapping.append(region)

        region_fields = [
            ("antiSMASH Region", "region"),
            ("antiSMASH Type", "type"),
            ("antiSMASH Region From", "start"),
            ("antiSMASH Region To", "end"),
            ("antiSMASH Similarity Confidence", "confidence"),
            ("antiSMASH Most similar known cluster", "known_cluster"),
        ]
        for output_col, source_key in region_fields:
            row_out[output_col] = merge_values([r.get(source_key, "") for r in overlapping])

        summaries = []
        for region in overlapping:
            parts = [text(region.get("region"))]
            if text(region.get("type")):
                parts.append(f"type={text(region.get('type'))}")
            parts.append(f"coordinates={region['start_num']}-{region['end_num']}")
            if text(region.get("confidence")):
                parts.append(f"confidence={text(region.get('confidence'))}")
            if text(region.get("known_cluster")):
                parts.append(f"most similar known cluster={text(region.get('known_cluster'))}")
            summaries.append("; ".join(parts))
        bgc_col = find_column(row_out.keys(), ["Associated with which BGC (based on antiSMASH result)"])
        if bgc_col:
            row_out[bgc_col] = " | ".join(summaries)

        basis = []
        if key in prokka_records:
            basis.append("Prokka CDS ID")
        if roary_matches:
            basis.append("Roary Prokka_GFF ID")
        if egg_matches:
            basis.append("eggNOG normalized locus tag")
        if has_kegg:
            basis.append("KEGG ID list")
        if overlapping:
            basis.append("antiSMASH coordinate overlap")
        row_out["Match basis"] = "; ".join(basis)
        row_out["CDS overlaps antiSMASH region"] = "yes" if overlapping else "no"

        output_rows.append(row_out)

    # Original columns first, followed by the joined source annotations.
    final_columns = list(base_columns)
    for col in DERIVED_COLUMNS:
        if col not in final_columns:
            final_columns.append(col)
    result = pd.DataFrame(output_rows)
    for col in final_columns:
        if col not in result.columns:
            result[col] = ""
    result = result[final_columns]

    # Use blank strings rather than NaN and keep a clean one-record-per-locus-tag output.
    if "Prokka Code" in result.columns:
        result = result.drop_duplicates(subset=["Prokka Code"], keep="first")
    result = result.fillna("")
    result.to_csv(output_path, sep="\t", index=False, encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Merge optional Bakta/Prokka, Roary, eggNOG, KEGG/BlastKOALA and antiSMASH files into one TSV row per CDS."
    )
    parser.add_argument("--version", action="version", version=f"AI-Parser {__version__}")
    parser.add_argument("--input", "-i", help="Legacy multi-sheet Excel workbook (alternative to individual source options)")
    parser.add_argument("--bakta", help="Bakta TSV/CSV/XLSX or GFF3 file (preferred master CDS list if supplied)")
    parser.add_argument("--prokka", help="Prokka TSV/CSV/XLSX or GFF3 file (master CDS list if Bakta is not supplied)")
    parser.add_argument("--roary", help="Roary gene_presence_absence.csv, TSV, or Excel export")
    parser.add_argument("--eggnog", help="eggNOG-mapper .emapper.annotations, TSV/CSV, or Excel export")
    parser.add_argument("--kegg", help="KEGG/BlastKOALA tab-delimited download, ID list, TSV/CSV, or Excel export")
    parser.add_argument("--antismash", help="antiSMASH region summary table in TSV/CSV/Excel format")
    parser.add_argument("--out", "-o", default="merged_gene_annotations.tsv", help="Output TSV path (default: merged_gene_annotations.tsv)")
    args = parser.parse_args()

    output_path = Path(args.out).expanduser().resolve()
    individual = {"bakta": args.bakta, "prokka": args.prokka, "roary": args.roary,
                  "eggnog": args.eggnog, "kegg": args.kegg, "antismash": args.antismash}
    if args.input and any(individual.values()):
        parser.error("Use either --input for a legacy workbook OR individual source options, not both.")
    if not args.input and not any(individual.values()):
        parser.error("Provide --bakta or --prokka as a master gene list, or use --input for a legacy workbook.")

    try:
        if args.input:
            input_path = Path(args.input).expanduser().resolve()
            if not input_path.is_file():
                parser.error(f"Input workbook does not exist: {input_path}")
            sheets = read_workbook_sheets(input_path)
        else:
            sheets = {}
            for kind, raw_path in individual.items():
                if raw_path:
                    path = Path(raw_path).expanduser().resolve()
                    # Bakta and Prokka share the same CDS parser; prefer Bakta when both supplied.
                    parse_kind = "prokka" if kind == "bakta" else kind
                    source_df = read_source_file(path, parse_kind)
                    if kind in {"bakta", "prokka"} and path.suffix.casefold() not in {".gff", ".gff3"}:
                        cols = {header_key(c) for c in (source_df.columns if source_df is not None else [])}
                        if not (cols & {"locustag", "prokkacode", "geneid", "id"}):
                            raise ValueError(
                                f"Cannot confidently identify {kind} input from its column headers: "
                                f"{list(source_df.columns)[:15] if source_df is not None else []}. "
                                "Please provide the native headered TSV/CSV or GFF3, or tell me the column layout."
                            )
                    sheets[parse_kind] = source_df
            if args.bakta:
                sheets["prokka"] = read_source_file(Path(args.bakta).expanduser().resolve(), "bakta")
                # parse_prokka reads the normalized Bakta header/feature table identically.
        # Validate identification rather than silently interpreting an unknown table as a source.
        for kind in ("prokka", "roary", "eggnog", "kegg", "antismash"):
            df = sheets.get(kind)
            if df is None or df.empty:
                continue
            cols = {header_key(c) for c in df.columns}
            if kind == "roary" and not (cols & {"gene", "annotation", "nonuniquegenename"}):
                raise ValueError(f"Cannot confidently identify {kind} input from its headers: {list(df.columns)[:12]}. Supply Roary gene_presence_absence.csv or an Excel export.")
            if kind == "eggnog" and not (cols & {"query", "locustag", "queryname", "proteinid", "#query"}):
                # #query is normalized to query by header_key; this branch covers recognized variants.
                raise ValueError(f"Cannot confidently identify eggNOG input from its headers: {list(df.columns)[:12]}. Supply eggNOG-mapper .emapper.annotations or an Excel export.")
            if kind == "antismash" and not (cols & {"from", "start", "regionfrom"} and cols & {"to", "end", "regionto"}):
                raise ValueError(f"Cannot confidently identify antiSMASH region table from its headers: {list(df.columns)[:12]}. Supply a region summary table with start/end coordinates.")
        result = build_output(sheets, output_path)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote: {output_path}")
    print(f"Genes/CDS rows: {len(result):,}")
    print(f"Columns: {len(result.columns):,}")
    print("Optional annotation inputs not supplied are skipped.")
    print("Note: KEGG/BlastKOALA assignments are imported only when the source file contains KO/definition columns.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
