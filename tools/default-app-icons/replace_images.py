import base64
import csv
import json
from pathlib import Path

BASE_DIR = Path(__file__).parent
CSV_FILE = BASE_DIR / "applist.csv"
TS_OUT = BASE_DIR.parent.parent / "src/webparts/myEnterpriseApps/assets/DefaultApps.ts"
SVG_DIR = BASE_DIR / "svg"

LABEL_COLUMN = "Label"
URL_COLUMN = "URL"
FALLBACK_ICON = (
    "data:image/svg+xml;base64,"
    "PHN2ZyBmaWxsPSJjdXJyZW50Q29sb3IiIGFyaWEtaGlkZGVuPSJ0cnVlIiB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNNC41IDE3QTEuNSAxLjUgMCAwIDEgMyAxNS42NVY0LjVjMC0uNzguNi0xLjQyIDEuMzYtMS41SDljLjc4IDAgMS40Mi42IDEuNSAxLjM2di40bDIuMTktMi4yN2ExLjUgMS41IDAgMCAxIDItLjE0bC4xMi4xIDIuNzYgMi43MmMuNTUuNTUuNiAxLjQyLjExIDIuMDJsLS4xLjExLTIuMzEgMi4yaC4yM2MuNzggMCAxLjQyLjYgMS41IDEuMzZ2NC42NGMwIC43OC0uNiAxLjQyLTEuMzYgMS41SDQuNVptNS02LjVINHY1YzAgLjIyLjE0LjQuMzMuNDdsLjA4LjAyLjA5LjAxaDV2LTUuNVptNiAwaC01VjE2aDVhLjUuNSAwIDAgMCAuNS0uNFYxMWEuNS41IDAgMCAwLS40MS0uNWgtLjA5Wm0tNS0yLjc5VjkuNWgxLjc5TDEwLjUgNy43MVpNOSA0LjAxSDQuNWEuNS41IDAgMCAwLS41LjR2NS4xaDUuNXYtNWEuNS41IDAgMCAwLS4zMy0uNDhsLS4wOC0uMDJIOVoiIGZpbGw9ImN1cnJlbnRDb2xvciI+PC9wYXRoPjwvc3ZnPg=="
)


def load_icon(label: str) -> str:
    """Return the SVG data URL for a label, or the fallback icon when missing."""
    svg_path = SVG_DIR / f"{label}.svg"
    if not svg_path.is_file():
        print(f"Waarschuwing: SVG ontbreekt, fallback gebruikt: {svg_path}")
        return FALLBACK_ICON

    encoded_svg = base64.b64encode(svg_path.read_bytes()).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded_svg}"


def is_empty_row(row: dict[str, str | None]) -> bool:
    """Return whether every value in a CSV row is empty."""
    return not any((value or "").strip() for value in row.values())


def load_apps() -> list[tuple[str, str, str]]:
    """Load and validate every non-empty CSV row as a default app."""
    with CSV_FILE.open(encoding="utf-8", newline="") as file:
        reader = csv.DictReader(file)
        rows = list(reader)
        fieldnames = reader.fieldnames or []

    missing_columns = [
        column for column in (LABEL_COLUMN, URL_COLUMN) if column not in fieldnames
    ]
    if missing_columns:
        raise ValueError(
            f"CSV mist verplichte kolom(men): {', '.join(missing_columns)}"
        )

    apps = []
    for line_number, row in enumerate(rows, start=2):
        if is_empty_row(row):
            continue

        name = (row.get(LABEL_COLUMN) or "").strip()
        url = (row.get(URL_COLUMN) or "").strip()
        missing_values = [
            column
            for column, value in ((LABEL_COLUMN, name), (URL_COLUMN, url))
            if not value
        ]
        if missing_values:
            raise ValueError(
                f"CSV-regel {line_number} mist verplichte waarde(n): "
                f"{', '.join(missing_values)}"
            )

        apps.append((name, url, load_icon(name)))

    return apps


def write_ts(apps: list[tuple[str, str, str]]) -> None:
    entries = []
    for name, url, icon in apps:
        entries.append(
            "  {\n"
            f"    name: {json.dumps(name)},\n"
            f"    url: {json.dumps(url)},\n"
            f"    icon: {json.dumps(icon)}\n"
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


def main() -> None:
    apps = load_apps()
    write_ts(apps)

    print("Klaar")
    print(f"   Apps: {len(apps)}")
    print(f"   TS: {TS_OUT.relative_to(BASE_DIR.parent.parent)}")


if __name__ == "__main__":
    main()
