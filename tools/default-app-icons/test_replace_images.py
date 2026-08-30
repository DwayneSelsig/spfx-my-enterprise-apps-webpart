import csv
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).parent))

import replace_images


class ReplaceImagesTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.base_dir = Path(self.temp_dir.name)
        self.csv_file = self.base_dir / "applist.csv"
        self.svg_dir = self.base_dir / "svg"
        self.ts_out = self.base_dir / "DefaultApps.ts"
        self.svg_dir.mkdir()
        self.paths = patch.multiple(
            replace_images,
            CSV_FILE=self.csv_file,
            SVG_DIR=self.svg_dir,
            TS_OUT=self.ts_out,
        )
        self.paths.start()

    def tearDown(self) -> None:
        self.paths.stop()
        self.temp_dir.cleanup()

    def write_csv(self, rows: list[dict[str, str]]) -> None:
        with self.csv_file.open("w", encoding="utf-8", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=["URL", "Label", "Image"])
            writer.writeheader()
            writer.writerows(rows)

    def write_svg(self, label: str) -> None:
        (self.svg_dir / f"{label}.svg").write_text("<svg />", encoding="utf-8")

    def test_generates_one_app_for_every_non_empty_csv_row(self) -> None:
        rows = [
            {"URL": "https://example.test/one", "Label": "One", "Image": "-"},
            {"URL": "", "Label": "", "Image": ""},
            {"URL": "https://example.test/two", "Label": "Two", "Image": "-"},
        ]
        self.write_csv(rows)
        self.write_svg("One")
        self.write_svg("Two")

        apps = replace_images.load_apps()
        replace_images.write_ts(apps)

        expected_count = sum(any(value.strip() for value in row.values()) for row in rows)
        generated = self.ts_out.read_text(encoding="utf-8")
        self.assertEqual(len(apps), expected_count)
        self.assertEqual(generated.count("    name: "), expected_count)
        self.assertEqual(generated.count("data:image/svg+xml;base64,"), expected_count)

    def test_uses_fallback_when_a_non_empty_row_has_no_svg(self) -> None:
        self.write_csv(
            [{"URL": "https://example.test/missing", "Label": "Missing", "Image": "-"}]
        )

        apps = replace_images.load_apps()
        replace_images.write_ts(apps)

        self.assertEqual(apps[0][2], replace_images.FALLBACK_ICON)
        self.assertIn(replace_images.FALLBACK_ICON, self.ts_out.read_text(encoding="utf-8"))

    def test_fails_when_a_non_empty_row_misses_a_required_value(self) -> None:
        self.write_csv([{"URL": "", "Label": "Missing URL", "Image": "-"}])
        self.write_svg("Missing URL")

        with self.assertRaisesRegex(ValueError, "URL"):
            replace_images.load_apps()


if __name__ == "__main__":
    unittest.main()
