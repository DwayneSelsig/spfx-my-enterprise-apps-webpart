import csv
import base64
import json
from pathlib import Path

BASE_DIR = Path(__file__).parent
CSV_FILE = BASE_DIR / "applist.csv"
TS_OUT = BASE_DIR.parent.parent / "src/webparts/myEnterpriseApps/assets/DefaultApps.ts"
SVG_DIR = BASE_DIR / "svg"

LABEL_COLUMN = "Label"
IMAGE_COLUMN = "Image"
URL_COLUMN = "URL"

def load_svg_base64(label: str):
    """
    Leest <Label>.svg en geeft base64 terug.
    Bestaat het bestand niet, dan None (rij blijft intact).
    """
    svg_path = SVG_DIR / f"{label}.svg"
    if not svg_path.exists():
        print(f"⚠️  SVG ontbreekt: {svg_path.name}")
        return None

    return base64.b64encode(svg_path.read_bytes()).decode("ascii")


def to_data_url(img: str) -> str:
    """Normalize base64 payload to a data URL."""
    if img.startswith("data:image"):
        return img
    return f"data:image/svg+xml;base64,{img}"


def write_ts(rows):
    entries = []
    for row in rows:
        name = (row.get(LABEL_COLUMN) or "").strip()
        url = (row.get(URL_COLUMN) or "").strip()
        img = (row.get(IMAGE_COLUMN) or "").strip()

        if not (name and url and img):
            continue

        entries.append(
            "  {\n"
            f"    name: {json.dumps(name)},\n"
            f"    url: {json.dumps(url)},\n"
            f"    icon: {json.dumps(to_data_url(img))}\n"
            "  },"
        )

    ts = "\n".join(
        [
            "// Auto-generated from tools/default-app-icons/applist.csv. Do not edit by hand.",
            "export interface IDefaultApp {",
            "  name: string;",
            "  url: string;",
            "  icon: string;",
            "}",
            "",
            "export const defaultApps: IDefaultApp[] = [",
            *entries,
            "];",
            "",
            "export const defaultAppsByName: Record<string, IDefaultApp> = defaultApps.reduce((acc, app) => {",
            "  acc[app.name.toLowerCase()] = app;",
            "  return acc;",
            "}, {} as Record<string, IDefaultApp>);",
            "",
        ]
    )

    TS_OUT.parent.mkdir(parents=True, exist_ok=True)
    TS_OUT.write_text(ts + "\n", encoding="utf-8")

def main():
    # CSV inlezen
    with CSV_FILE.open(encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        fieldnames = reader.fieldnames or []

    if (
        LABEL_COLUMN not in fieldnames
        or IMAGE_COLUMN not in fieldnames
        or URL_COLUMN not in fieldnames
    ):
        raise ValueError(
            f"CSV moet kolommen '{URL_COLUMN}', '{LABEL_COLUMN}' en '{IMAGE_COLUMN}' bevatten"
        )

    replaced = 0
    skipped = 0

    # Image vervangen waar mogelijk
    for row in rows:
        label = (row.get(LABEL_COLUMN) or "").strip()
        if not label:
            skipped += 1
            continue

        encoded_svg = load_svg_base64(label)
        if encoded_svg is not None:
            row[IMAGE_COLUMN] = encoded_svg
            replaced += 1
        else:
            skipped += 1

        # CSV overschrijven
    #with CSV_FILE.open("w", encoding="utf-8", newline="") as f:
    #    writer = csv.DictWriter(
    #        f,
    #        fieldnames=fieldnames,
    #        quoting=csv.QUOTE_MINIMAL
    #    )
    #   writer.writeheader()
    #   writer.writerows(rows)

    write_ts(rows)

    print("✅ Klaar")
    print(f"   Vervangen: {replaced}")
    print(f"   Overgeslagen: {skipped}")
    print(f"   TS: {TS_OUT.relative_to(BASE_DIR.parent.parent)}")

if __name__ == "__main__":
    main()
