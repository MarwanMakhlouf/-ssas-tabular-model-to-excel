import argparse
import json
from pathlib import Path
import sys
import pandas as pd


def main():
    parser = argparse.ArgumentParser(
        description="Extract table/column names and translations from a JSON file and save as Excel sheets."
    )
    parser.add_argument("input", help="Path to the input JSON file")
    parser.add_argument("outdir", help="Output directory for Excel files")
    args = parser.parse_args()

    input_path = Path(args.input)
    outdir = Path(args.outdir)

    if not input_path.is_file():
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        sys.exit(2)

    try:
        outdir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"Error: could not create output directory {outdir}: {e}", file=sys.stderr)
        sys.exit(3)

    with input_path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    table_translations = []
    column_translations = []
    measure_translations = []
    relationships = []
    table_data = []

    tables = data["model"]["tables"]
    for table in tables:
        isCalculatedTable = False
        partition = table.get("partitions", [])
        if partition and partition[0]["source"]["type"] == "calculated":
            isCalculatedTable = True
        for column in table.get("columns", []):
            isCalculatedColumn = False
            if "type" in column and column["type"] == "calculated":
                isCalculatedColumn = True
            table_data.append(
                [
                    table["name"],
                    "Calculated" if isCalculatedTable else "",
                    partition[0]["source"]["expression"] if isCalculatedTable else "",
                    column["name"],
                    "Column",
                    "Calculated" if isCalculatedColumn else "",
                    column["expression"] if isCalculatedColumn else "",
                ]
            )
        if "measures" in table:
            for measure in table["measures"]:
                table_data.append(
                    [
                        table["name"],
                        "Calculated" if isCalculatedTable else "",
                        (partition[0]["source"]["expression"] if isCalculatedTable else ""),
                        measure["name"],
                        "Measure",
                        "Calculated",
                        measure["expression"],
                    ]
                )

    relationship_data = data["model"]["relationships"]
    for rel in relationship_data:
        relationships.append([rel["fromTable"], rel["fromColumn"], rel["toTable"], rel["toColumn"]])

    translation_data = data["model"]["cultures"][0]["translations"]["model"]["tables"]
    for table in translation_data:
        table_translations.append([table["name"], table.get("translatedCaption", "")])
        for column in table.get("columns", []):
            column_translations.append([column["name"], column.get("translatedCaption", "")])
        for measure in table.get("measures", []):
            measure_translations.append([measure["name"], measure.get("translatedCaption", "")])

    df_relationships = pd.DataFrame(
        relationships,
        columns=["From Table", "From Column", "To Table", "To Column"],
    )

    df_table_translations = pd.DataFrame(
        table_translations,
        columns=["Table Name", "Translation"],
    )

    df_column_translations = pd.DataFrame(
        column_translations,
        columns=["Column Name", "Translation"],
    )

    df_measure_translations = pd.DataFrame(
        measure_translations,
        columns=["Measure Name", "Translation"],
    )

    df_table_data = pd.DataFrame(
        table_data,
        columns=[
            "Table Name",
            "Table Type",
            "Table Expression",
            "Column/Measure Name",
            "Column/Measure",
            "Column/Measure Type",
            "Column/Measure Expression",
        ],
    )

    df_relationships.to_excel(outdir / "relationships.xlsx", index=False)
    df_table_translations.to_excel(outdir / "table_translations.xlsx", index=False)
    df_column_translations.to_excel(outdir / "column_translations.xlsx", index=False)
    df_measure_translations.to_excel(outdir / "measure_translations.xlsx", index=False)
    df_table_data.to_excel(outdir / "table_data.xlsx", index=False)


if __name__ == "__main__":
    main()
