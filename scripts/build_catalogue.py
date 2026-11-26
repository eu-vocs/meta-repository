import argparse
import csv
import html
import re
from collections import defaultdict
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import Request, urlopen

from openpyxl import load_workbook


RESOURCE_SHEET = "Data_resources"
USAGE_SHEET = "DC_specific_data_usage"

RESOURCE_COLUMNS = [
    "Dataset_ID",
    "Short_name",
    "Full_name",
    "Project_generated",
    "URL",
    "Data_types",
    "Data_publicity",
    "Notes",
]

USAGE_COLUMNS = [
    "Subproject id",
    "Dataset_ID",
    "Relationship",
]

# Optional HTML table fields shown by default.
# Use --fields= to show no optional fields.
DEFAULT_FIELDS = ["Notes"]

# Data_types and Data_publicity have their own dedicated columns.
UNAVAILABLE_OPTIONAL_FIELDS = {
    "Dataset_ID",
    "Project_generated",
    "URL",
    "Data_types",
    "Data_publicity",
    "Short_name",
}


def text(value):
    """Convert an Excel cell value to trimmed text."""
    if value is None:
        return ""
    return str(value).strip()


def normalize_id(value):
    """Normalize IDs for matching without displaying them."""
    return text(value).casefold()


def normalize_label(value):
    """Normalize a label for matching, ignoring spaces and underscores."""
    return re.sub(r"[\s_-]+", "", text(value).casefold())


def check_columns(sheet, expected_columns):
    """Check that a worksheet contains the required column headers."""
    headers = {
        text(cell.value): cell.column
        for cell in sheet[1]
        if text(cell.value)
    }

    missing = [name for name in expected_columns if name not in headers]
    if missing:
        raise ValueError(
            f"Sheet '{sheet.title}' is missing columns: "
            + ", ".join(missing)
        )

    return headers


def rows_from_sheet(sheet, expected_columns):
    """Read worksheet rows as dictionaries using the header names."""
    headers = check_columns(sheet, expected_columns)

    for row_number in range(2, sheet.max_row + 1):
        row = {
            name: sheet.cell(row=row_number, column=headers[name]).value
            for name in expected_columns
        }

        if any(text(value) for value in row.values()):
            yield row_number, row


def safe_url(value):
    """Accept an HTTP(S) URL or DOI and return a safe clickable URL."""
    value = text(value)
    if not value:
        return ""

    value = re.sub(r"^doi:\s*", "", value, flags=re.IGNORECASE)

    if re.match(r"^10\.\d{4,9}/\S+$", value, flags=re.IGNORECASE):
        value = "https://doi.org/" + value
    elif value.lower().startswith("doi.org/"):
        value = "https://" + value

    parts = urlsplit(value)
    if parts.scheme.lower() not in {"http", "https"} or not parts.netloc:
        return ""

    return value


def check_link(url, timeout=8):
    """
    Check whether a link responds.

    Some servers reject HEAD requests, so try GET for common method errors.
    """
    headers = {"User-Agent": "ProjectDataCatalogueLinkChecker/1.0"}

    for method in ("HEAD", "GET"):
        request = Request(url, headers=headers, method=method)

        try:
            with urlopen(request, timeout=timeout) as response:
                status = response.status
                if 200 <= status < 400:
                    return True, f"HTTP {status}"
                return False, f"HTTP {status}"

        except HTTPError as error:
            if method == "HEAD" and error.code in {403, 405, 501}:
                continue
            return False, f"HTTP {error.code}"

        except (URLError, TimeoutError, OSError) as error:
            return False, str(error)

    return False, "Could not verify the link"


def split_tags(value):
    """Split a semicolon-separated cell into non-empty labels."""
    return [part.strip() for part in text(value).split(";") if part.strip()]


def data_type_style(label):
    """Return the icon and CSS class for a data-type label."""
    label_key = normalize_label(label)

    styles = {
        "audio": ("graphic_eq", "type-audio"),
        "image": ("image", "type-image"),
        "video": ("videocam", "type-video"),
        "behavioralresponses": ("psychology", "type-behaviour"),
        "fmri": ("neurology", "type-fmri"),
        "other": ("dataset", "type-other"),
    }

    return styles.get(label_key, ("help", "type-unknown"))


def publicity_style(value):
    """Return the icon and CSS class for a publicity label."""
    label_key = normalize_label(value)

    styles = {
        "public": ("lock_open", "publicity-public"),
        "restrictedpublic": ("lock", "publicity-restricted"),
        "purchaseable": ("shopping_cart", "publicity-purchaseable"),
        "unclear": ("help", "publicity-unclear"),
    }

    return styles.get(label_key, ("help", "publicity-unknown"))


def render_badges(value, field_name, use_colors):
    """Render icon-and-text badges for data types or publicity labels."""
    labels = split_tags(value)

    if not labels:
        return '<span class="muted">Not specified</span>'

    badges = []

    for label in labels:
        if field_name == "Data_types":
            icon_name, color_class = data_type_style(label)
        else:
            icon_name, color_class = publicity_style(label)

        if not use_colors:
            color_class = ""

        class_names = ["badge"]
        if color_class:
            class_names.append(color_class)

        badge = (
            '<span class="'
            + " ".join(class_names)
            + '">'
            '<span class="material-symbols-outlined" aria-hidden="true">'
            + html.escape(icon_name)
            + "</span>"
            "<span>"
            + html.escape(label)
            + "</span>"
            "</span>"
        )
        badges.append(badge)

    return '<div class="badges">' + "".join(badges) + "</div>"


def is_project_generated(value):
    """Recognize common affirmative values in Project_generated."""
    affirmative_values = {
        "yes",
        "y",
        "true",
        "1",
        "projectgenerated",
        "generated",
        "collected",
    }
    return normalize_label(value) in affirmative_values


def project_role_html(dataset, usage):
    """Show a prominent generated badge or subdued existing-data badge."""
    if is_project_generated(dataset["Project_generated"]):
        return (
            '<span class="role-badge role-new">'
            '<span class="material-symbols-outlined" aria-hidden="true">'
            "auto_awesome"
            "</span>"
            "<span>New! Project-generated</span>"
            "</span>"
        )

    relationship = text(usage["relationship"])

    existing_badge = (
        '<span class="role-badge role-existing">'
        '<span class="material-symbols-outlined" aria-hidden="true">'
        "database"
        "</span>"
        "<span>Existing</span>"
        "</span>"
    )

    if relationship:
        return existing_badge + (
            '<div class="relationship-detail">'
            + html.escape(relationship)
            + "</div>"
        )

    return existing_badge


def dataset_title(dataset):
    return (
        text(dataset["Short_name"])
        or text(dataset["Full_name"])
        or "Dataset"
    )


def render_optional_field(field_name, dataset):
    """Render an optional table field as escaped text."""
    value = text(dataset.get(field_name))
    if not value:
        return '<span class="muted">—</span>'
    return html.escape(value)


def build_rows(datasets):
    """Create one row per subproject–dataset relationship."""
    rows = []

    for dataset in datasets:
        usages = dataset["usages"]

        if not usages:
            usages = [{"subproject": "Not linked", "relationship": ""}]

        for usage in usages:
            rows.append((dataset, usage))

    rows.sort(
        key=lambda item: (
            item[1]["subproject"].casefold(),
            dataset_title(item[0]).casefold(),
        )
    )

    return rows


def render_table(rows, fields, publicity_colors, data_type_colors):
    headers = [
        "Subproject",
        "Dataset",
        "Project role",
        "Data types",
        "Data publicity",
    ]

    for field in fields:
        headers.append(field.replace("_", " "))

    header_cells = []

    for label in headers:
        header_cells.append(
            '<th scope="col">'
            '<button type="button" class="sort-button">'
            + html.escape(label)
            + ' <span aria-hidden="true">↕</span>'
            + "</button></th>"
        )

    body_rows = []

    for dataset, usage in rows:
        title = dataset_title(dataset)
        full_name = text(dataset["Full_name"])
        url = safe_url(dataset["URL"])

        title_html = html.escape(title)

        if url:
            escaped_url = html.escape(url, quote=True)
            title_html = (
                '<a href="'
                + escaped_url
                + '" target="_blank" rel="noopener noreferrer">'
                + title_html
                + "</a>"
            )

        if full_name and full_name.casefold() != title.casefold():
            title_html += (
                '<div class="full-name">'
                + html.escape(full_name)
                + "</div>"
            )

        subproject = text(usage["subproject"]) or "Not specified"

        cells = [
            "<td>" + html.escape(subproject) + "</td>",
            "<td>" + title_html + "</td>",
            "<td>" + project_role_html(dataset, usage) + "</td>",
            "<td>"
            + render_badges(
                dataset["Data_types"],
                "Data_types",
                data_type_colors,
            )
            + "</td>",
            "<td>"
            + render_badges(
                dataset["Data_publicity"],
                "Data_publicity",
                publicity_colors,
            )
            + "</td>",
        ]

        for field in fields:
            cells.append(
                "<td>" + render_optional_field(field, dataset) + "</td>"
            )

        searchable_values = [
            subproject,
            text(usage["relationship"]),
            title,
            full_name,
            text(dataset["Project_generated"]),
            text(dataset["Data_types"]),
            text(dataset["Data_publicity"]),
        ]

        for field in fields:
            searchable_values.append(text(dataset[field]))

        searchable = " ".join(searchable_values).casefold()
        escaped_search = html.escape(searchable, quote=True)

        body_rows.append(
            '<tr data-search="'
            + escaped_search
            + '">'
            + "".join(cells)
            + "</tr>"
        )

    if not body_rows:
        return "<p>No dataset entries to display.</p>"

    return (
        '<div class="table-wrap">'
        '<table id="catalogue">'
        "<thead><tr>"
        + "".join(header_cells)
        + "</tr></thead>"
        "<tbody>"
        + "".join(body_rows)
        + "</tbody>"
        "</table>"
        "</div>"
    )


def legend_item(icon_name, label, css_class=""):
    """Render one icon-and-text legend entry."""
    class_names = ["legend-item"]
    if css_class:
        class_names.append(css_class)

    return (
        '<span class="'
        + " ".join(class_names)
        + '">'
        '<span class="material-symbols-outlined" aria-hidden="true">'
        + html.escape(icon_name)
        + "</span>"
        "<span>"
        + html.escape(label)
        + "</span>"
        "</span>"
    )


def build_legend(publicity_colors, data_type_colors):
    """Build legends for data types, publicity, and project role."""
    data_type_items = [
        ("graphic_eq", "Audio", "type-audio"),
        ("image", "Image", "type-image"),
        ("videocam", "Video", "type-video"),
        ("psychology", "Behavioral responses", "type-behaviour"),
        ("neurology", "fMRI", "type-fmri"),
        ("dataset", "Other", "type-other"),
    ]

    publicity_items = [
        ("lock_open", "Public", "publicity-public"),
        ("lock", "Restricted public", "publicity-restricted"),
        ("shopping_cart", "Purchaseable", "publicity-purchaseable"),
        ("help", "UNCLEAR — check status", "publicity-unclear"),
    ]

    data_type_html = []
    for icon_name, label, css_class in data_type_items:
        if not data_type_colors:
            css_class = ""
        data_type_html.append(legend_item(icon_name, label, css_class))

    publicity_html = []
    for icon_name, label, css_class in publicity_items:
        if not publicity_colors:
            css_class = ""
        publicity_html.append(legend_item(icon_name, label, css_class))

    project_role_html = [
        legend_item("auto_awesome", "New! Project-generated", "role-new"),
        legend_item("database", "Existing dataset", "role-existing"),
    ]

    return (
        '<section class="legend" aria-label="Catalogue icon legend">'
        "<h2>Legend</h2>"
        "<h3>Data types</h3>"
        '<div class="legend-items">'
        + "".join(data_type_html)
        + "</div>"
        "<h3>Data publicity</h3>"
        '<div class="legend-items">'
        + "".join(publicity_html)
        + "</div>"
        "<h3>Project role</h3>"
        '<div class="legend-items">'
        + "".join(project_role_html)
        + "</div>"
        "</section>"
    )


def build_html(
    rows,
    fields,
    source_name,
    publicity_colors,
    data_type_colors,
):
    table_html = render_table(
        rows,
        fields,
        publicity_colors,
        data_type_colors,
    )
    legend_html = build_legend(publicity_colors, data_type_colors)

    template = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Project data catalogue</title>

  <link rel="stylesheet"
        href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined">

  <style>
    :root {
      font-family: system-ui, sans-serif;
      line-height: 1.5;
      color: #202124;
      background: #f4f6f8;
    }
    body { max-width: 1400px; margin: 0 auto; padding: 1.5rem; }
    header, .legend {
      background: white;
      border-radius: 0.6rem;
      padding: 1.25rem;
    }
    h1 { margin-top: 0; }
    input {
      box-sizing: border-box;
      width: 100%;
      max-width: 38rem;
      padding: 0.7rem;
      margin-top: 0.5rem;
    }
    .table-wrap { overflow-x: auto; margin-top: 1rem; }
    table { width: 100%; border-collapse: collapse; background: white; }
    th, td {
      text-align: left;
      vertical-align: top;
      padding: 0.75rem;
      border: 1px solid #d7dce2;
    }
    th { background: #eaf0f6; position: sticky; top: 0; }
    tbody tr:nth-child(even) { background: #f8fafc; }
    .sort-button {
      border: 0;
      background: transparent;
      color: inherit;
      font: inherit;
      font-weight: 700;
      padding: 0;
      cursor: pointer;
    }
    .sort-button:focus, input:focus, a:focus {
      outline: 3px solid #f0b429;
      outline-offset: 2px;
    }
    .badges, .legend-items {
      display: flex;
      flex-wrap: wrap;
      gap: 0.4rem;
    }
    .badge, .legend-item, .role-badge {
      display: inline-flex;
      align-items: center;
      gap: 0.3rem;
      border-radius: 1rem;
      padding: 0.2rem 0.6rem;
    }
    .badge, .legend-item { background: #eef1f4; }
    .role-badge { white-space: nowrap; }
    .role-new {
      background: #fff0c2;
      border: 1px solid #d49a00;
      color: #493700;
      font-weight: 700;
    }
    .role-existing {
      background: #eceff1;
      border: 1px solid #c3cbd1;
      color: #424b50;
    }
    .relationship-detail {
      margin-top: 0.2rem;
      color: #555;
      font-size: 0.9rem;
    }
    .type-audio { background: #dceeff; color: #163c5a; }
    .type-image { background: #eee5fb; color: #452b67; }
    .type-video { background: #fbe4e8; color: #642b37; }
    .type-behaviour { background: #e1f1e5; color: #234b2c; }
    .type-fmri { background: #d9f0ed; color: #174b46; }
    .type-other { background: #e8ebef; color: #333b43; }
    .type-unknown { background: #eeeeee; color: #333333; }
    .publicity-public { background: #176b3a; color: #ffffff; }
    .publicity-restricted { background: #ffedaa; color: #433600; }
    .publicity-purchaseable { background: #ffedaa; color: #433600; }
    .publicity-unclear, .publicity-unknown {
      background: #f4d6d6;
      color: #761f1f;
      border: 1px solid #bd5b5b;
    }
    .material-symbols-outlined {
      font-family: "Material Symbols Outlined";
      font-weight: normal;
      font-style: normal;
      font-size: 1.2rem;
      line-height: 1;
      letter-spacing: normal;
      text-transform: none;
      display: inline-block;
      white-space: nowrap;
      direction: ltr;
      font-feature-settings: "liga";
      -webkit-font-feature-settings: "liga";
      -webkit-font-smoothing: antialiased;
    }
    .full-name, .muted { color: #555; }
    .full-name { font-size: 0.9rem; }
    .legend { margin-top: 1rem; }
    .legend h2 { margin-top: 0; }
    .legend h3 { font-size: 1rem; margin-bottom: 0.4rem; }
    .legend-items { margin-bottom: 0.8rem; }
    a { color: #0645ad; }
  </style>
</head>
<body>
  <header>
    <h1>Project data catalogue</h1>
    <p>Each row links a dataset to a subproject. A dataset associated with
       multiple subprojects appears once per subproject relationship.</p>
    <label for="search">Search the catalogue</label><br>
    <input id="search" type="search"
           placeholder="Search subprojects, datasets, or data types">
    <p>Click a column heading to sort the table.</p>
  </header>

  __LEGEND__

  <main>
    __TABLE__
  </main>

  <script>
    const searchInput = document.getElementById("search");
    const table = document.getElementById("catalogue");

    if (searchInput && table) {
      searchInput.addEventListener("input", () => {
        const query = searchInput.value.trim().toLowerCase();
        for (const row of table.tBodies[0].rows) {
          row.hidden = !row.dataset.search.includes(query);
        }
      });

      for (const header of table.tHead.rows[0].cells) {
        const button = header.querySelector("button");
        button.addEventListener("click", () => {
          const columnIndex = header.cellIndex;
          const ascending = header.dataset.direction !== "ascending";
          header.dataset.direction = ascending ? "ascending" : "descending";

          const rows = Array.from(table.tBodies[0].rows);
          rows.sort((firstRow, secondRow) => {
            const first = firstRow.cells[columnIndex].innerText.trim();
            const second = secondRow.cells[columnIndex].innerText.trim();
            const comparison = first.localeCompare(second, undefined, {
              numeric: true,
              sensitivity: "base"
            });
            return ascending ? comparison : -comparison;
          });

          for (const row of rows) {
            table.tBodies[0].appendChild(row);
          }
        });
      }
    }
  </script>
</body>
</html>
"""

    return (
        template
        .replace("__LEGEND__", legend_html)
        .replace("__TABLE__", table_html)
    )


def main():
    parser = argparse.ArgumentParser(
        description="Generate a CSV catalogue and sortable HTML from Excel."
    )

    parser.add_argument(
        "workbook",
        help="Path to the source .xlsx workbook",
    )
    parser.add_argument(
        "--output-dir",
        default="build",
        help="Output folder (default: build)",
    )
    parser.add_argument(
        "--fields",
        default=",".join(DEFAULT_FIELDS),
        help=(
            "Comma-separated optional Data_resources fields to display. "
            "Use --fields= to display none."
        ),
    )
    parser.add_argument(
        "--check-links",
        action="store_true",
        help="Check URLs/DOIs online and warn about failures.",
    )
    parser.add_argument(
        "--public-only",
        action="store_true",
        help="Include only rows matching the selected publicity labels.",
    )
    parser.add_argument(
        "--publicity-values",
        default="Public",
        help=(
            "Comma-separated Data_publicity labels allowed with "
            "--public-only. Default: Public"
        ),
    )

    publicity_group = parser.add_mutually_exclusive_group()
    publicity_group.add_argument(
        "--publicity-colors",
        dest="publicity_colors",
        action="store_true",
        help="Enable publicity-status colours (enabled by default).",
    )
    publicity_group.add_argument(
        "--no-publicity-colors",
        dest="publicity_colors",
        action="store_false",
        help="Disable publicity-status colours.",
    )
    parser.set_defaults(publicity_colors=True)

    data_type_group = parser.add_mutually_exclusive_group()
    data_type_group.add_argument(
        "--data-type-colors",
        dest="data_type_colors",
        action="store_true",
        help="Enable data-type colours (enabled by default).",
    )
    data_type_group.add_argument(
        "--no-data-type-colors",
        dest="data_type_colors",
        action="store_false",
        help="Disable data-type colours.",
    )
    parser.set_defaults(data_type_colors=True)

    args = parser.parse_args()

    workbook_path = Path(args.workbook)
    output_dir = Path(args.output_dir)

    if not workbook_path.exists():
        raise SystemExit(f"Workbook not found: {workbook_path}")

    fields = [
        field.strip()
        for field in args.fields.split(",")
        if field.strip()
    ]

    available_optional_fields = [
        column
        for column in RESOURCE_COLUMNS
        if column not in UNAVAILABLE_OPTIONAL_FIELDS
    ]

    invalid_fields = [
        field for field in fields
        if field not in available_optional_fields
    ]

    if invalid_fields:
        raise SystemExit(
            "Invalid --fields value(s): "
            + ", ".join(invalid_fields)
            + ". Available optional fields: "
            + ", ".join(available_optional_fields)
        )

    workbook = load_workbook(
        workbook_path,
        read_only=True,
        data_only=True,
    )

    for sheet_name in (RESOURCE_SHEET, USAGE_SHEET):
        if sheet_name not in workbook.sheetnames:
            raise SystemExit(
                f"Missing sheet '{sheet_name}'. "
                f"Found: {', '.join(workbook.sheetnames)}"
            )

    usages_by_dataset = defaultdict(list)

    for row_number, row in rows_from_sheet(
        workbook[USAGE_SHEET],
        USAGE_COLUMNS,
    ):
        dataset_key = normalize_id(row["Dataset_ID"])

        if not dataset_key:
            print(
                f"Warning: usage row {row_number} has no Dataset_ID; skipped."
            )
            continue

        usages_by_dataset[dataset_key].append(
            {
                "subproject": text(row["Subproject id"]),
                "relationship": text(row["Relationship"]),
            }
        )

    allowed_publicity = {
        value.strip().casefold()
        for value in args.publicity_values.split(",")
        if value.strip()
    }

    datasets = []
    known_ids = set()

    for row_number, row in rows_from_sheet(
        workbook[RESOURCE_SHEET],
        RESOURCE_COLUMNS,
    ):
        dataset_key = normalize_id(row["Dataset_ID"])

        if not dataset_key:
            print(
                f"Warning: resource row {row_number} has no Dataset_ID; skipped."
            )
            continue

        if dataset_key in known_ids:
            print(
                f"Warning: duplicate Dataset_ID "
                f"'{text(row['Dataset_ID'])}' at row {row_number}."
            )

        known_ids.add(dataset_key)

        publicity = text(row["Data_publicity"]).casefold()
        if args.public_only and publicity not in allowed_publicity:
            continue

        dataset = {
            column: row[column]
            for column in RESOURCE_COLUMNS
        }
        dataset["usages"] = usages_by_dataset.get(dataset_key, [])
        datasets.append(dataset)

    unmatched_ids = set(usages_by_dataset) - known_ids

    for unmatched_id in sorted(unmatched_ids):
        print(
            f"Warning: usage sheet refers to Dataset_ID '{unmatched_id}', "
            "which is absent from Data_resources."
        )

    if args.check_links:
        checked_urls = {}

        for dataset in datasets:
            raw_value = text(dataset["URL"])
            if not raw_value:
                continue

            url = safe_url(raw_value)
            name = dataset_title(dataset)

            if not url:
                print(
                    f"WARNING: '{name}' has an invalid URL/DOI: {raw_value}"
                )
                continue

            if url not in checked_urls:
                checked_urls[url] = check_link(url)

            success, message = checked_urls[url]

            if success:
                print(f"Link OK: {raw_value} ({message})")
            else:
                print(
                    f"WARNING: link check failed for '{name}': "
                    f"{raw_value} ({message})"
                )

    rows = build_rows(datasets)
    output_dir.mkdir(parents=True, exist_ok=True)

    csv_path = output_dir / "catalogue.csv"
    html_path = output_dir / "index.html"

    # Dataset_ID is used for joining, but is omitted from the output.
    csv_columns = [
        "Short_name",
        "Full_name",
        "URL",
        "Project role",
        "Data_types",
        "Data_publicity",
        "Notes",
        "Subproject id",
    ]

    with csv_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=csv_columns)
        writer.writeheader()

        for dataset, usage in rows:
            if is_project_generated(dataset["Project_generated"]):
                project_role = "New! Project-generated"
            else:
                relationship = text(usage["relationship"])
                project_role = "Existing"
                if relationship:
                    project_role += " — " + relationship

            output_row = {
                "Short_name": text(dataset["Short_name"]),
                "Full_name": text(dataset["Full_name"]),
                "URL": text(dataset["URL"]),
                "Project role": project_role,
                "Data_types": text(dataset["Data_types"]),
                "Data_publicity": text(dataset["Data_publicity"]),
                "Notes": text(dataset["Notes"]),
                "Subproject id": text(usage["subproject"]),
            }
            writer.writerow(output_row)

    html_path.write_text(
        build_html(
            rows,
            fields,
            workbook_path.name,
            args.publicity_colors,
            args.data_type_colors,
        ),
        encoding="utf-8",
    )

    print(f"Generated {len(rows)} subproject-dataset row(s).")
    print(f"CSV:  {csv_path.resolve()}")
    print(f"HTML: {html_path.resolve()}")

    if args.public_only:
        print(
            "Included publicity labels: "
            + ", ".join(sorted(allowed_publicity))
        )


if __name__ == "__main__":
    main()
