import csv
import re
import argparse
from pathlib import Path


def parse_bookmark_line(line):
    line = line.lstrip("- ").strip()

    match = re.match(r"(.+?)\s*-\s*(https?://\S+)\s*-\s*(.+)", line)
    if match:
        return match.group(1).strip(), match.group(2).strip(), match.group(3).strip()

    match = re.match(r"(.+?)\s*-\s*(https?://\S+)", line)
    if match:
        return match.group(1).strip(), match.group(2).strip(), ""

    match = re.match(r"(.+?)\s*:\s*(https?://\S+)", line)
    if match:
        return match.group(1).strip(), match.group(2).strip(), ""

    match = re.match(r"(https?://\S+)", line)
    if match:
        return "", match.group(1).strip(), ""

    return "", "", line


def md_to_csv(input_path, output_path):
    rows = []
    current_category = ""

    with open(input_path, "r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()

            if not line:
                continue

            if line.startswith("**") and line.endswith("**"):
                current_category = line.strip("*")
                continue

            if line.startswith("- "):
                name, url, description = parse_bookmark_line(line)
                rows.append({
                    "category": current_category,
                    "name": name,
                    "url": url,
                    "description": description
                })

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(
            csvfile,
            fieldnames=["category", "name", "url", "description"]
        )
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(
        description="Convert a Markdown bookmark list into CSV"
    )
    parser.add_argument("input_md", help="Path to input .md file")
    parser.add_argument("output_csv", help="Path to output .csv file")

    args = parser.parse_args()

    input_path = Path(args.input_md)
    output_path = Path(args.output_csv)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    md_to_csv(input_path, output_path)
    print(f"Converted {input_path} → {output_path}")


if __name__ == "__main__":
    main()
