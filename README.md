# MDB Studio

MDB Studio is a Windows desktop database editor currently distributed through standard numbered GitHub releases. The application is actively developed and does not yet replace every Microsoft Access feature. It aims to support Access-like database management for .mdb and .accdb files. This repository contains transparent **C# / WPF** source and a Windows build workflow.

## Current features

Open existing Microsoft Access files; browse all records through pages of 1,000, 5,000 or 10,000 rows, or optionally load **All records**; show the actual table record count and page numbers; edit, insert and delete records in tables with primary keys; save changes transactionally with a backup of the original file; run read-only SELECT queries without a fixed 2,000-result ceiling; export **all saved table records** to CSV. The Check updates button reads public release metadata from the official GitHub repository and, after user confirmation, can open a release in the default browser. It never downloads or executes update binaries.

Not yet implemented: full Access table design, relationships editing, forms, reports, saved query designer, macros, VBA, full fidelity with all Access data types, or support for every Access database.

## Safety and compatibility

**Do not use the application on irreplaceable data. Start with a separate database copy.** MDB Studio does not convert the original file. It uses Windows Microsoft ACE OLE DB. The matching 64-bit Microsoft 365 Access Runtime / ACE provider must be installed. The app detects ACE 16.0 and 12.0 and reports provider errors.

The earlier experimental **Go-based** EXE distributed outside this repository was flagged by Microsoft Defender. Do not run, restore, whitelist, or redistribute those earlier EXEs. This is a fresh Windows C# implementation, not a copy of those binaries. A successful GitHub build does **not** establish that the resulting executable has passed antivirus scanning or is code-signed. Do not bypass Windows Defender warnings. Standard GitHub releases reflect the user's successful functional testing, not independent antivirus certification.

## New interface and clipboard records

The redesigned blue-and-white Windows desktop interface includes the MDB Studio geometric database logo as the app icon and header artwork, a structured toolbar, searchable table navigation, tabbed table/SQL workspace, contextual right-click record actions, and a status bar. Use **Ctrl+C** or Copy records to copy selected full rows as tab-separated values with a header for spreadsheet compatibility. Use **Ctrl+V** or Paste records to add new rows from MDB Studio or a spreadsheet. Duplicate repeats selected rows only when Access generates the primary key automatically; tables with manual keys require new unique IDs. Delete marks selected rows for deletion only after confirmation. Pasted, duplicated, and deleted rows remain pending until **Save Changes**. Copying from a read-only table works, but pasting and deleting require a detectable primary key. Clipboard pastes are capped at 10,000 rows per operation.

**Important:** Data type validation, NOT NULL requirements, and relationship constraints remain enforced by Access during saving. Work on a disposable database copy and inspect the pending edits. Design mockups are not evidence of runtime testing.

## Large-table browsing and updates

MDB Studio no longer truncates table browsing at the first 2,000 records. The default page is 5,000 rows. Use **Next**, **Previous**, and the page-size selector to browse every row, or choose **All records** if the table comfortably fits in memory. Paging relies on the Access OLE DB adapter's row-window support rather than an Access SQL `OFFSET` clause. Tables with detectable primary keys are sorted by their key for consistent page boundaries; tables without keys are read-only, and their page order can shift if another program changes the database while browsing. Switching pages asks about unsaved edits so they cannot be silently lost. Table export streams the full saved table, not only the page.

**Check updates** uses the GitHub Releases API for `TheImposter1000/MDBStudio`, compares semantic versions including preview numbers, and offers only to open the official release page. The application does not self-update, auto-download, execute binaries, bypass antivirus warnings, or store release executables.

## Password-protected databases

If an Access MDB/ACCDB file reports "Not a valid password", MDB Studio asks for a database password that you already know. It passes the password to the Microsoft ACE provider using `Jet OLEDB:Database Password` and does not write it to a file or application logs. This feature does not remove, crack, discover, or bypass unknown passwords. Unsupported or damaged database formats can also give authentication-like errors.

**TEW IX:** The game protects its working `TEW9.mdb` and `TEW9Save.mdb` files and does not publish their database password. Use the game's **External Editor** function to produce an editable, password-free MDB intended for editing and subsequent in-game import. The game's **Create MDB File** option is for external viewing only and cannot be imported as a playable database. Do not attempt direct writes to protected game databases. See the TEW IX Player's Handbook and Grey Dog Software developer notes.

## Build

GitHub Actions runs the Windows workflow in .github/workflows/windows-build.yml. It restores, compiles, and publishes the self-contained x64 application. The release workflow publishes normal version tags such as `v1.0.0` and `v1.0.1`. Each release is a portable ZIP containing a multi-file application folder, not an installer. Keep all published files together when testing. The project targets .NET 10.

## Roadmap

Database structure/design tools, referential integrity and relationships, visual query builder, forms, reports, import tools, complex field types, performance improvements for very large record sets, automated tests and MDB/ACCDB fixtures, Windows runtime tests, security review, packaging, signed release, and updater.

The project is not affiliated with Microsoft.

## Versioning

The repository uses normal GitHub releases for versions such as `v1.0.0`, `v1.0.1` and `v1.1.0`. Before publishing, update `MDBStudio.csproj` and `UpdateChecker.CurrentReleaseTag` to exactly the same version, then run the release workflow. Previews are optional and are not the default distribution channel. The **Check & install updates** button offers a confirmed, verified in-place installation of newer releases.

## Paste and duplicate improvements (v1.0.1)

Copying complete selected rows and pasting them into the same table now creates **new records** with fresh identifiers, instead of failing when the source records' primary keys already exist. MDB Studio checks primary keys against all saved records in the table, not just the current page, and accounts for pending rows. For a single manually assigned numeric primary key, it allocates IDs above the current global maximum. Single GUID keys receive new GUIDs. Text primary keys can receive new numbered suffixes.

If a table has a text `Recordname` column, existing names such as `Prairie Cup013` are advanced to the next available name (for example `Prairie Cup024`) while retaining their prefix and padding width. The other record fields are copied unchanged and should be reviewed. This does **not** automatically modify `BeltOrder` or other domain-specific sequence fields. Pasting into a table with composite keys or unsupported identity types that cannot be safely regenerated still requires manual IDs, and database unique indexes are checked by Access on save. The right-click context menu retains a previously selected group of rows. Copied and pasted rows remain pending until **Save Changes**, which creates a backup before writing.

This release uses the standard GitHub version format, not a preview.

## In-app updates (v1.0.2)

The Check & install updates button downloads new releases directly from the official GitHub repository. The app validates a published SHA256 manifest, refuses to discard unsaved database edits, and launches an independent updater from a temporary folder. The updater waits for MDB Studio to close, verifies integrity again, extracts the release, backs up the program files it replaces, installs the update, and relaunches the application. If installation fails, it attempts to restore the previous version. No MDB/ACCDB database files or user-created files are part of the release archive. No antivirus exceptions or elevation bypasses are made.

One-time setup: MDB Studio v1.0.1 and earlier do not contain the updater. You will need to extract v1.0.2 once into your existing writable program folder. Every subsequent release containing the updater can be installed directly from the application's button without manually downloading a ZIP. The application is not code-signed or independently antivirus certified. Do not bypass Windows Defender warnings.

## Cell-based clipboard and Fill Down (v1.0.3)

MDB Studio now supports spreadsheet-style cell selection, not just whole-row copying. Select one cell and press **Ctrl+C** to copy its value only, then select a range of cells in the same or other data column and press **Ctrl+V** to repeat that value into every selected cell. Select a rectangular block of cells to copy and paste a matching block, or paste a TSV block from Excel into a single starting cell or a matching multiple-sized selected rectangle. **Ctrl+D** fills selected column cells downward using the value in the topmost selected cell in each column. **Delete** clears selected cells where their column permits blank values; the explicit Delete rows button only acts on selected record headers. Copy and Paste toolbar buttons and the right-click menu use cell operations instead of creating new records. Record count and other columns remain unchanged. Non-writable, generated and primary key columns are protected. Data type and length validation occur before making any changes, with a confirmation for more than 1,000 cells. **Save Changes** still commits edits to Access after backup.

For large tables, select the cells on the current loaded page, use page-size controls or All records when appropriate. Clipboard operations cannot silently edit unloaded rows. Old record duplication helpers remain in source for compatibility tests but are not bound to Ctrl+C / Ctrl+V.

## Restore deleting records (v1.0.4)

Whole-record deletion remains available with cell-based copy and paste. Select one or several records using **any cells in the relevant rows** or their **row headers**, then click **Delete records** in the toolbar or context menu. **Ctrl+Delete** performs the same record-deletion action. The application asks for confirmation and marks only the selected records for deletion. **Save Changes** commits deletions with the existing MDB backup safeguard. **Delete** alone clears selected cell values, rather than deleting whole rows, unless complete rows were selected using the row headers. Selected rows are deduplicated, and unsaved new records can be removed without issuing a database DELETE. Tables without a detectable primary key remain read-only. If no rows are selected, the Delete records command remains unavailable.

## Excel import and query export (v1.1.0)

Run a SELECT query in the SQL Query tab and click **Export results to Excel (.xlsx)**. MDB Studio creates an actual Excel workbook containing the returned query columns and records, with frozen headers, an autofilter and readable formatting. This exports the query results rather than an entire database table. Excel row limits still apply.

Click **Import Excel** to select an `.xlsx` file, choose a worksheet, review its columns and first 30 records, and either **create a new MDB table** or **append rows to an existing table**. New tables use an AutoNumber primary key named `_MDBStudioID`; the imported columns are stored as long-text fields to preserve formatted identifiers and leading zeros. Existing-table imports match Excel header names to database columns, skip AutoNumber fields, validate data types, and are subject to primary key, relationship and unique-index restrictions. Unmatched headings are ignored. Use one sheet per import; files larger than 50 MB or worksheets above 100,000 records / 200 columns must be split.

Excel imports are **confirmed and committed separately** from Save Changes, after creating a database backup. This is deliberate: you see a preview and approve the destination before MDB Studio changes the database. Pending edits in the current table must be saved first. The importer does not execute macros and supports `.xlsx`, not legacy `.xls` or `.xlsm`. Importing into a new table is a new MDB table operation, not a new MDB file.

The Excel integration uses the open-source ClosedXML 0.105.1 library. As always, experiment on a copied database before using valuable data.

## CSV import and query export (v1.1.1)

The **Import Excel/CSV** toolbar command now accepts comma-separated `.csv` files and `.xlsx` workbooks. When importing a CSV, the first record supplies field names, and the same preview and destination wizard lets you either create a new MDB table or append rows to an existing table. CSV values are treated as text until matched with the destination MDB field types, preserving identifiers such as `000012`. It supports standard double-quoted fields with embedded commas, quotes, and newlines as well as UTF-8 input including BOM. The importer rejects mismatched field counts and enforces a limit of 50 MB, 100,000 data records, and 200 columns. Imports require confirmation and a database backup; they commit separately from Save Changes.

Under **SQL Query**, run a SELECT query and choose either **Export results to Excel (.xlsx)** or **Export results to CSV (.csv)**. CSV exports contain all returned query rows and headings, use UTF-8 with BOM, and properly escape commas, quotes and line breaks. When exporting text values that spreadsheet software might treat as formulas (beginning with `=`, `+`, `-` or `@`), MDB Studio prefixes an apostrophe in the CSV to prevent formula execution. This intentionally changes that cell's CSV representation and should be accounted for if round-tripping formula-like text. Numeric negative values are exported unchanged.

Existing full-table CSV export remains available. This update does not add legacy `.xls` support. Use copies of important databases when validating imports.


## Importing Access Yes/No flags (v1.1.2)

Excel and CSV imports into existing MDB tables now understand native Access Yes/No values. For Boolean database columns, -1, 1, True, Yes, Y, On, and T import as True. The values 0, False, No, N, Off, and F import as False, regardless of letter case; surrounding whitespace is permitted. A value of -1 remains numeric when the destination field is an integer, rather than a Yes/No field. Invalid Boolean values such as 2 or -2 still stop the import with a field-specific error before the database is modified. Empty cells continue to follow the destination field's nullability constraints. Bulk cell pasting uses the same Access-compatible Boolean parser.

This fixes the error reported when an import contains -1 in a Yes/No field such as Victor. The change does not relax primary-key, unique-index, or relationship constraints; those are still enforced by Access. As before, import is confirmed and a backup is created prior to writing. Test updates against a copy of the database.

## Verify the installed application and diagnose updates (v1.1.3)

Earlier releases displayed an incorrect hardcoded `MDB Studio v1.1.1` badge even after a newer version was installed. From v1.1.3, the badge and window title come from the executable's actual assembly version. Select **Installation details** next to **Check & install updates** to see the running version, executable path, installation folder, updater presence and log location. The details are copied to the clipboard when requested. If two extracted copies exist, this reveals which executable is running.

The updater now writes an installation and recovery log to `%LOCALAPPDATA%\MDBStudio\update.log`, including installation and restart status and errors. It does not record database contents or passwords. The app checks that it can write to its installation folder before closing to install an update. Windows-protected folders may require moving the portable app to a writable location. No antivirus bypass or automatic privilege elevation is attempted.

v1.1.2 already included the fix to import native Access Yes/No values such as `-1`. The earlier static version badge could not establish whether that version actually installed. If the import still fails, use **Installation details** to verify the running executable and its location. Retest on a copy of the MDB database and do not bypass antivirus warnings.

## TEW angle winner sentinel protection (v1.1.4)

The `Victor` field in show history can hold a numeric worker winner ID for a match, or the numeric sentinel `-1` for an angle with no match winner. This is not the same semantic meaning as an Access Yes/No Boolean value. MDB Studio preserves `-1` as a signed integer when the existing destination field is numeric, whether `Int16` or `Int32`. Match winner IDs and other numeric values are also preserved.

For an existing destination field named `Victor` that is actually defined as Yes/No (Boolean), MDB Studio rejects the angle sentinel with an actionable schema warning rather than silently converting it into `True`. A Yes/No field cannot faithfully store a numeric no-winner sentinel; select the correct existing numeric field or table. Other genuine Access Yes/No fields still accept `-1` as True and `0` as False.

Regression tests cover CSV imports mixing angle and match rows, preservation of numeric Victor -1 and winner IDs, type mismatch rejection, and regular Access Boolean flags. This is a data integrity protection and does not change any original MDB table schema. As with all TEW database changes, test imports on a copy before importing into the working database.

## Spreadsheet numeric import reliability (v1.1.5)

MDB Studio now accepts numeric cells such as TEW `Victor = -1` even if previous CSV exports stored them as `'-1`, if a pasted value uses a Unicode minus glyph, or if a number includes a common invisible leading/trailing mark. Cleanup happens **only on numeric destination fields**, leaving original text fields and Recordname values unchanged. Invalid or ambiguous inputs still fail before importing any records, now with the raw Unicode code points (such as U+2212) and normalized value in the diagnostic message.

CSV export no longer prepends apostrophes to valid signed numeric **text literals**, such as `-1`. Spreadsheet formula injection protection still applies to expressions such as `=HYPERLINK(...)` and negative formulas such as `-SUM(...)`. Files exported by earlier MDB Studio versions that added an apostrophe to a negative numeric string can now be reimported into a numeric column. Genuine Access Yes/No fields still accept -1 as True; TEW's numeric Victor=-1 stays -1.

Regression coverage includes Excel/CSV-style values, legacy quote-prefixed numbers, UTF-8/Unicode minus signs, numeric IDs, preserved text fields, malicious formula-like strings, error diagnostics, updater checksum validation and rollback. The actual user's spreadsheet has not been supplied for byte-by-byte verification, so test on a copy of the MDB database first.


## Universal record deletion (v1.1.6)

Delete Records is available for every opened table, including tables without a primary key. Select any cell or row header, use Delete Records or Ctrl+Delete, and confirm. Deletions stay pending until Save Changes; MDB Studio backs up the database before committing them. On keyless tables, normal cell editing and row insertion remain disabled.

For a table without a primary key, MDB Studio builds a parameterized identity predicate from supported short text and scalar fields, including nullable values. Before each deletion, within the same transaction, it counts every stored record matching that predicate and requires that all matching records were explicitly selected. This prevents accidentally deleting additional identical records, including those on another page. If no safe predicate exists, or a matching duplicate was not selected, the entire save is rolled back and an explanatory error is shown. Long text and binary fields are excluded from identity predicates. No Access table schema is changed. Database constraints and relationships may still block deletion. Test on a backup of important TEW data.


## Table-header filters (v1.1.7)

MDB Studio now offers full-table column filtering above the records grid. Click any column header to choose that field, or choose it in **Filter column**. Pick a condition (Contains, Starts with, Ends with, Equals, Does not equal, Is blank, Is not blank, and supported number/date comparisons). Select a suggested distinct value or type your own, then choose **Apply**. Additional filters on different columns combine with AND. **Remove selected** removes the active filter for the chosen column; **Clear all** restores the unfiltered table.

Unlike the separate **Search displayed rows** text box, these column filters are evaluated by Microsoft ACE against the entire saved MDB table before pagination. Record counts and Next/Previous navigation use the filtered result, rather than searching only the visible page. Up to 250 distinct values are suggested from a supported column; for large text/complex fields, manually type a filter value. Filter expressions are parameterized, preserve numeric identifiers and TEW negative sentinel values, and do not modify MDB data. Applying or clearing filters checks for unsaved edits first and never discards them without confirmation.

Filtering is separate from permanently sorting or grouping company records. The field names come from the currently opened table; MDB Studio does not infer foreign-key relationships or translate company IDs to display names automatically.


## New database and double-click file opening (v1.1.8)

Use **New database** to create a blank MDB (Jet 4) or ACCDB (Access 2007+) file. MDB Studio uses DAO or ADOX from the 64-bit Access Database Engine, validates the created file with ACE, and refuses to overwrite any existing file. It builds the new database at a temporary path, verifies it, then moves it into place. If DAO/ADOX are missing, creation fails with an explanatory message. Microsoft Access desktop is not required. A blank database contains no tables and is not a TEW IX game save; use **Import Excel/CSV** to create its first table.

Double clicking a database in Windows Explorer now opens that file inside MDB Studio when it is selected as Windows' default handler. Use **Open with setup** to register MDB Studio as a supported application for MDB and ACCDB under the current user's account. Then select MDB Studio in Windows **Default Apps** or use **Open with** and choose **Always**. The registration does not change existing defaults automatically. If the app is moved to a different folder, run Open with setup again.

Tests cover new file path validation, non-overwrite behavior, startup arguments, shell command quoting and existing application regressions. Actual database file creation requires the local Microsoft ACE/DAO runtime and has not been exercised on a Windows machine with that component.
