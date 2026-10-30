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

The repository uses normal GitHub releases for versions such as `v1.0.0`, `v1.0.1` and `v1.1.0`. Before publishing, update `MDBStudio.csproj` and `UpdateChecker.CurrentReleaseTag` to exactly the same version, then run the release workflow. Previews are optional and are not the default distribution channel. The **Check updates** button opens the official release page; it does not automatically install anything.

## Paste and duplicate improvements (v1.0.1)

Copying complete selected rows and pasting them into the same table now creates **new records** with fresh identifiers, instead of failing when the source records' primary keys already exist. MDB Studio checks primary keys against all saved records in the table, not just the current page, and accounts for pending rows. For a single manually assigned numeric primary key, it allocates IDs above the current global maximum. Single GUID keys receive new GUIDs. Text primary keys can receive new numbered suffixes.

If a table has a text `Recordname` column, existing names such as `Prairie Cup013` are advanced to the next available name (for example `Prairie Cup024`) while retaining their prefix and padding width. The other record fields are copied unchanged and should be reviewed. This does **not** automatically modify `BeltOrder` or other domain-specific sequence fields. Pasting into a table with composite keys or unsupported identity types that cannot be safely regenerated still requires manual IDs, and database unique indexes are checked by Access on save. The right-click context menu retains a previously selected group of rows. Copied and pasted rows remain pending until **Save Changes**, which creates a backup before writing.

This release uses the standard GitHub version format, not a preview.
