#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple
import importlib.util
import subprocess


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


print('''
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

Example (Windows Command Prompt):
    py ai-parser.py --input "Sample data.xlsx" --out merged_gene_annotations.tsv

https://github.com/zblara/ai-parser.git
©2026 Zachary B. Lara 
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


def read_workbook_sheets(path: Path) -> Dict[str, pd.DataFrame]:
    """Read sheets with per-sheet header settings, including the headerless sample tabs."""
    xls = pd.ExcelFile(path, engine="openpyxl")
    available = {name.casefold(): name for name in xls.sheet_names}
    sheets: Dict[str, pd.DataFrame] = {}

    # These sheets have column headers in row 1.
    for wanted in ("Output", "Roary", "eggnog", "antiSMASH"):
        real_name = available.get(wanted.casefold())
        if real_name:
            df = pd.read_excel(xls, sheet_name=real_name, header=0, dtype=object)
            df.columns = [text(c) for c in df.columns]
            df = df.dropna(how="all")
            sheets[wanted.casefold()] = df

    # Prokka in the supplied workbook is a headerless 7-column extract.
    real_name = available.get("prokka")
    if real_name:
        df = pd.read_excel(xls, sheet_name=real_name, header=None, dtype=object)
        df = df.dropna(how="all").reset_index(drop=True)
        sheets["prokka"] = df

    # KEGG in the sample is an ID-only, headerless single column.
    real_name = available.get("kegg")
    if real_name:
        df = pd.read_excel(xls, sheet_name=real_name, header=None, dtype=object)
        df = df.dropna(how="all").reset_index(drop=True)
        sheets["kegg"] = df

    return sheets


def parse_prokka(df: Optional[pd.DataFrame]) -> Dict[str, Dict[str, Any]]:
    """Return records keyed by normalized locus tag; supports headerless sample format."""
    records: Dict[str, Dict[str, Any]] = {}
    if df is None or df.empty:
        return records

    first_row = [text(v) for v in df.iloc[0].tolist()]
    first_keys = {header_key(v) for v in first_row}
    headered = bool(first_keys & {"locustag", "prokkacode", "ftype", "seqid"})

    if headered:
        headers = first_row
        data = df.iloc[1:].copy()
        data.columns = headers
        locus_col = find_column(data.columns, ["locus_tag", "Prokka Code", "gene_id", "ID"])
        contig_col = find_column(data.columns, ["contig", "seqid", "sequence", "accession"])
        feature_col = find_column(data.columns, ["ftype", "feature", "type"])
        start_col = find_column(data.columns, ["start", "begin"])
        end_col = find_column(data.columns, ["end", "stop"])
        strand_col = find_column(data.columns, ["strand"])
        product_col = find_column(data.columns, ["product", "annotation", "description"])
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


def parse_kegg(df: Optional[pd.DataFrame]) -> Tuple[set, bool]:
    """Return normalized IDs and whether non-ID annotation columns were supplied."""
    ids = set()
    if df is None or df.empty:
        return ids, False

    # Use the first column as the ID list in the sample workbook. Header aliases are skipped.
    id_headers = {"id", "geneid", "locustag", "prokkacode", "query", "keggid", "gene"}
    for value in df.iloc[:, 0].tolist():
        raw = text(value)
        if not raw or header_key(raw) in id_headers:
            continue
        key = normalise_locus_tag(raw)
        if key:
            ids.add(key)

    annotations_available = False
    if len(df.columns) > 1:
        for col in df.columns[1:]:
            if any(not is_blank(v) for v in df[col].tolist()):
                annotations_available = True
                break
    return ids, annotations_available


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
    kegg_ids, kegg_annotations_available = parse_kegg(sheets.get("kegg"))
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
        row_out["KEGG ID list match"] = "yes" if has_kegg else "no"
        row_out["KEGG annotation availability"] = (
            "Annotation columns supplied; inspect original KEGG sheet"
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
        description="Merge Roary, Prokka, eggNOG, KEGG and antiSMASH annotations into one TSV row per CDS."
    )
    parser.add_argument("--input", "-i", required=True, help="Input Excel workbook containing the source sheets")
    parser.add_argument("--out", "-o", default="merged_gene_annotations.tsv", help="Output TSV path (default: merged_gene_annotations.tsv)")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.out).expanduser().resolve()
    if not input_path.is_file():
        parser.error(f"Input workbook does not exist: {input_path}")

    try:
        sheets = read_workbook_sheets(input_path)
        result = build_output(sheets, output_path)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote: {output_path}")
    print(f"Genes/CDS rows: {len(result):,}")
    print(f"Columns: {len(result.columns):,}")
    print("Note: if the KEGG sheet only contains IDs, the script can report ID matches but cannot infer KO or pathway annotations.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
