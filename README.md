# MDF2Excel

Convert a Microsoft SQL Server `.mdf` database file into a single Excel (`.xlsx`)
workbook — **without installing SQL Server**.

![License: MIT](https://img.shields.io/badge/license-MIT-yellow)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

MDF2Excel reads the SQL Server on-disk page format directly, so you can recover
the schema and every row of data from a detached, orphaned, or legacy database
file — even when the original SQL Server instance no longer exists.

---

## The problem it solves

Older databases are hard to read once their server is gone:

- **Newer SQL Server versions refuse to attach SQL Server 2000 / 2005 files.**
  The on-disk format changed, so you can't simply "open" an old `.mdf`.
- Recovering data usually means finding a compatible SQL Server version,
  restoring backups, and exporting table by table — slow and technical.
- Many schools, shops, and small businesses still hold years of records in
  these files with no DBA and no SQL Server license available.

MDF2Excel skips all of that. Point it at an `.mdf` file and it exports every
user table into one tidy Excel workbook you can open anywhere.

## Who this helps

- **Schools / institutes** migrating off legacy student-management systems
- **Accountants & businesses** recovering invoices, ledgers, and registers
- **IT support & developers** extracting data from client backups
- **Anyone** handed a `.mdf` file with no idea what server created it

## Supported SQL Server versions

| SQL Server       | Internal version |
|------------------|------------------|
| SQL Server 2000  | 539              |
| SQL Server 2005  | 611 / 612        |
| SQL Server 2008  | 655              |
| SQL Server 2008 R2 | 661            |

## Features

- **No SQL Server required** — parses the binary file directly
- **One workbook, every table** — each table becomes a worksheet
- **`_Summary` sheet** — table names with row and column counts
- **Schema preserved** — column names and types are read from the catalog
- **Correct data types** — `datetime`, `smalldatetime`, `decimal`/`numeric`
  (scale applied), `money`, `int`, `bit`, text, and more
- **Read-only** — the source file is never modified
- **Cross-platform** — Windows, macOS, and Linux
- **GUI + command line** — friendly window or scriptable CLI

## Quick start

### Windows (standalone app)

1. Download `MDF2Excel.exe` from the
   [Releases](https://github.com/ovaisi/mdf2excel/releases) page.
2. Double-click it (no Python needed).
3. Click **Browse...**, choose your `.mdf` file, then **Convert to Excel**.

You can also drag an `.mdf` file onto the executable to prefill the input.

### macOS

```bash
chmod +x mdf2excel.command mdf2excel.sh
./mdf2excel.command
```

The launcher installs the one dependency (`openpyxl`) automatically if needed.

### Linux

```bash
chmod +x mdf2excel.sh
./mdf2excel.sh                 # graphical window
sudo apt install python3-tk    # only needed for the GUI
```

### Python (any OS)

```bash
pip install openpyxl
python mdf2excel.py "database.mdf"
```

## Command-line reference

```bash
# Convert a file (output defaults to "<name>_AllTables.xlsx")
python mdf2excel.py "database.mdf"

# Choose the output path
python mdf2excel.py "database.mdf" -o "export.xlsx"

# Just inspect, without converting
python mdf2excel.py "database.mdf" --info
python mdf2excel.py "database.mdf" --tables
```

The same arguments work with `mdf2excel.sh` and `mdf2excel.command`.

## Output format

| Sheet          | Contents                                                        |
|----------------|-----------------------------------------------------------------|
| `_Summary`     | Every table with its row count, column count, and status        |
| one per table  | The table's columns as the header row, followed by its data rows|

Dates are written as `YYYY-MM-DD HH:MM:SS` text; decimals have their scale
applied; `NULL` cells are left empty.

## How it works

1. **Boot page** — reads the database name and version from page 9.
2. **System catalog** — reads `sysobjects` / `syscolumns` / `sysindexes`
   (SQL 2000) or `sysschobjs` / `syscolpars` (SQL 2005+) to discover tables,
   columns, types, and page locations.
3. **Data pages** — walks each table's pages (IAM chains and data pages),
   decoding records into typed values.
4. **Excel** — writes every table into one `.xlsx` workbook via `openpyxl`.

## Limitations

- Off-row LOB data (`text` / `ntext` / `image` stored on separate pages) is
  reported as a pointer like `<LOB:16bytes>` rather than reassembled.
- Read-only; it never alters the source `.mdf`.

## Project structure

```
MDF2Excel.exe           standalone Windows app (GUI)
mdf2excel.py            the tool — GUI + command line
mdf_parser.py           low-level MDF page/record parser
mdf2excel.sh            macOS / Linux launcher
mdf2excel.command       macOS double-click launcher
build-all-platforms.yml  GitHub Actions CI template (native builds)
```

## Building native binaries

`build-all-platforms.yml` is a ready-to-use GitHub Actions workflow. Copy it
to `.github/workflows/build.yml` in the repository, then run it from the
**Actions** tab. It builds native executables for Windows, macOS, and Linux,
downloadable as artifacts.

## Contributing

Bug reports, feature requests, and pull requests are welcome. Please open an
issue first to discuss larger changes.

## License

Released under the [MIT License](LICENSE).

## Acknowledgments

Based on [`pfak/mssql-mdf-parser`](https://github.com/pfak/mssql-mdf-parser)
(MIT), with fixes for SQL Server 2000 record decoding, date/time field order,
and decimal precision/scale handling.

## Author & support

**Muhammad Ovais Jahanzaib** — software developer at Vergemobile.

- GitHub: [@ovaisi](https://github.com/ovaisi)
- LinkedIn: [ovaisjanzeb](https://www.linkedin.com/in/ovaisjanzeb/)

If this tool rescued an old database or saved you a few hours, you can support
its development here:

[☕ Support the project](https://janzeb.gumroad.com/l/sjkzgm)
