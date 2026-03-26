# MsAccessExtract

A .NET 10 console application that extracts all objects from a Microsoft Access database (`.accdb` / `.mdb`) into organized text files, enabling **Git version control** for Access applications.

Inspired by [MSAccessVCS](https://github.com/joyfullservice/msaccess-vcs-addin), but as a **standalone executable** — no add-in installation required.

---

## Extracted Objects

| Object | Method | Format | Folder |
|---|---|---|---|
| VBA Modules | `VBComponent.Export()` | `.bas` | `modules/` |
| VBA Classes | `VBComponent.Export()` | `.cls` | `classes/` |
| Forms | `SaveAsText` + sanitization | `.bas` | `forms/` |
| Reports | `SaveAsText` + sanitization | `.bas` | `reports/` |
| Queries | `SaveAsText` + SQL via DAO | `.bas` + `.sql` | `queries/` |
| Macros | `SaveAsText` + sanitization | `.bas` | `macros/` |
| Tables | DAO `TableDefs` | `.json` | `tables/` |
| Relationships | DAO `Relations` | `relations.json` | `relations/` |
| VBA References | `VBProject.References` | `database-properties.json` | root |

## Sanitization (Clean Diffs)

`SaveAsText` exports contain noise that generates false diffs in Git (checksums, GUIDs, printer settings). MsAccessExtract automatically removes:

- `Checksum`, `NoSaveCTID`, `GUID`
- `PrtMip`, `PrtDevMode`, `PrtDevNames` (printer settings)
- `NameMap` (binary map)
- OLE blobs (`dbLongBinary "OLE"`)

---

## Prerequisites

- **Microsoft Access** installed (365, 2021, 2019, or 2016)
- **"Trust access to the VBA project object model"** enabled:
  - Access → File → Options → Trust Center → Trust Center Settings
  - Macro Settings → ✅ Trust access to the VBA project object model

> **Note:** Use the **x64** build for 64-bit Office and **x86** for 32-bit Office.

---

## Usage

### Auto-detect mode (recommended)

Place the executable in the same folder as your Access database and run it:

```
📁 MyFolder/
├── MyDatabase.accdb          ← your database
└── MsAccessExtract.exe       ← the tool
```

```powershell
.\MsAccessExtract.exe
```

### Specify a file path

```powershell
.\MsAccessExtract.exe C:\path\to\MyDatabase.accdb
```

### Output

```
📁 MyDatabase_src/
├── forms/
│   ├── frmCustomers.bas
│   └── frmProducts.bas
├── reports/
│   └── rptSales.bas
├── queries/
│   ├── qryActiveCustomers.bas
│   └── qryActiveCustomers.sql
├── modules/
│   └── modUtilities.bas
├── classes/
│   └── clsLogger.cls
├── macros/
│   └── AutoExec.bas
├── tables/
│   ├── Customers.json
│   └── Products.json
├── relations/
│   └── relations.json
└── database-properties.json
```

### Git Workflow

```bash
cd MyDatabase_src
git init
git add .
git commit -m "Initial export from Access"
```

On subsequent exports, just run the tool again and Git will show only the real changes.

---

## Build

### Build Requirements

- [.NET 10 SDK](https://dotnet.microsoft.com/download) (or later)

### Compile

```powershell
dotnet build -c Release
```

### Publish (single-file, self-contained)

```powershell
# 64-bit Office
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o ./publish/x64

# 32-bit Office
dotnet publish -c Release -r win-x86 --self-contained -p:PublishSingleFile=true -o ./publish/x86
```

---

## Project Structure

```
MsAccessExtract/
├── Program.cs                      # Entry point + CLI
├── AccessExtractor.cs              # Main orchestrator
├── Extractors/
│   ├── ModuleExtractor.cs          # VBA modules and classes
│   ├── FormExtractor.cs            # Forms
│   ├── ReportExtractor.cs          # Reports
│   ├── QueryExtractor.cs           # Queries
│   ├── MacroExtractor.cs           # Macros
│   ├── TableExtractor.cs           # Tables (JSON)
│   ├── RelationshipExtractor.cs    # Relationships (JSON)
│   └── ReferenceExtractor.cs       # VBA references
├── Sanitizers/
│   └── SaveAsTextSanitizer.cs      # Removes noise from exports
├── Models/
│   ├── TableSchema.cs
│   ├── RelationshipInfo.cs
│   └── VbaReference.cs
├── Helpers/
│   ├── AccessConstants.cs          # COM constants (late binding)
│   ├── ComHelper.cs                # COM lifecycle management
│   └── ConsoleLogger.cs            # Colored logger
└── MsAccessExtract.csproj
```

---

## Exit Codes

| Code | Meaning |
|---|---|
| `0` | Extraction completed successfully |
| `1` | Fatal error (database not found, Access not installed, etc.) |
| `2` | Extraction completed with partial errors |

---

## License

MIT
