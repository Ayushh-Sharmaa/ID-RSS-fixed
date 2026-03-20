import re
import os
from docx import Document


def extract_field(text, label):
    """
    Finds 'Label- Value' anywhere in text.
    Captures everything after the dash until end of line.
    Works for multi-word values like full names.
    """
    escaped = re.escape(label)
    # Match label followed by dash (with optional spaces) then capture rest of line
    pattern = rf"{escaped}\s*-\s*(.+?)(?:\r?\n|$)"
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return "—"


def extract_from_docx(filepath, fields):
    """
    Reads a .docx file and extracts the requested fields.
    fields: list of label strings e.g. ["Name", "Serial No."]
    Returns a dict with 'file' key plus one key per field.
    """
    try:
        doc = Document(filepath)
        # Join all paragraphs with newlines so regex can find fields anywhere
        text = "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        # If file can't be read, return empty result
        result = {"file": os.path.basename(filepath), "error": str(e)}
        for f in fields:
            result[f] = "—"
        return result

    result = {"file": os.path.basename(filepath)}

    found_count = 0
    for field in fields:
        value = extract_field(text, field)
        result[field] = value
        if value != "—":
            found_count += 1

    # Confidence = how many fields were actually found
    result["confidence"] = round(found_count / len(fields), 2) if fields else 0.0

    return result


def process_folder(folder_path, fields):
    """
    Processes all .docx files in a folder.
    folder_path: string path to the folder
    fields: list of field label strings to extract
    Returns list of dicts, one per file.
    """
    results = []

    if not os.path.exists(folder_path):
        return []

    files = sorted([
        f for f in os.listdir(folder_path)
        if f.endswith(".docx")
    ])

    for fname in files:
        filepath = os.path.join(folder_path, fname)
        result = extract_from_docx(filepath, fields)
        results.append(result)

    return results


# ── Quick test ────────────────────────────────────────────────
if __name__ == "__main__":
    import sys

    folder = "demo_files"
    fields = ["Name", "Serial No."]

    print(f"Scanning: {folder}")
    print(f"Fields:   {fields}\n")

    results = process_folder(folder, fields)

    for r in results:
        print(f"{r['file']:<25}", end="")
        for f in fields:
            print(f"  {f}: {r[f]:<30}", end="")
        print(f"  confidence: {r['confidence']}")

    found = sum(1 for r in results if r["confidence"] == 1.0)
    print(f"\n{found}/{len(results)} files fully extracted")