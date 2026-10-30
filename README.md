# MDB Studio

MDB Studio is an experimental Windows desktop database editor. It aims to support Access-like database management for .mdb and .accdb files. This repository contains transparent **C# / WPF** source and a Windows build workflow.

## Current preview scope

Open existing Microsoft Access files; browse tables; view up to 2,000 rows per table; edit, insert and delete records when a primary key is available; save changes transactionally with a backup of the original file; run read-only SELECT queries; export the currently loaded table as CSV.

Not yet implemented: full Access table design, relationships editing, forms, reports, saved query designer, macros, VBA, full fidelity with all Access data types, or support for every Access database.

## Safety and compatibility

**Do not use this preview on irreplaceable data. Start with a separate database copy.** MDB Studio does not convert the original file. It uses Windows Microsoft ACE OLE DB. The matching 64-bit Microsoft 365 Access Runtime / ACE provider must be installed. The app detects ACE 16.0 and 12.0 and reports provider errors.

The earlier experimental **Go-based** EXE distributed outside this repository was flagged by Microsoft Defender. Do not run, restore, whitelist, or redistribute those earlier EXEs. This is a fresh Windows C# implementation, not a copy of those binaries. A successful GitHub build does **not** establish that the resulting executable has passed antivirus scanning, is code-signed, or is production-safe. The GitHub artifact is explicitly labeled UNVERIFIED.

## Build

GitHub Actions runs the Windows workflow in .github/workflows/windows-build.yml. It restores, compiles, and publishes the self-contained x64 application. The resulting build artifact is a multi-file folder (not a packed or obfuscated EXE). Keep all published files together when testing. The project targets .NET 10.

## Roadmap

Database structure/design tools, referential integrity and relationships, visual query builder, forms, reports, import tools, complex field types, performance/paging, test MDB/ACCDB fixtures, Windows runtime tests, security review, packaging, signed release, and updater.

The project is not affiliated with Microsoft.
