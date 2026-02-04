# Power BI Schema Extraction Tools

This repository contains tools and scripts to extract schema from Power BI Desktop models via the XMLA endpoint.

## Files

### Documentation
- **[tabular-editor-skill.md](tabular-editor-skill.md)** - Comprehensive guide on connecting to Power BI Desktop through XMLA and extracting schema

### Scripts

1. **[Extract-PBISchema.ps1](Extract-PBISchema.ps1)** - PowerShell script
   - Full-featured extraction with auto-detection
   - Exports to multiple formats (JSON, TXT, DAX, CSV)
   - Detailed error handling and reporting

2. **[ExtractSchema.csx](ExtractSchema.csx)** - C# script for Tabular Editor
   - Run directly from Tabular Editor
   - Timestamp-based file naming
   - Separates measures, calculated columns, and tables

3. **[extract_pbi_schema.py](extract_pbi_schema.py)** - Python script
   - Cross-platform support
   - Command-line interface
   - JSON-based metadata export

## Quick Start

### PowerShell
```powershell
.\Extract-PBISchema.ps1 -ServerName "localhost:PORT" -DatabaseName "YourModel"
```

### Python
```bash
pip install pythonnet
python extract_pbi_schema.py --server localhost:PORT --database "YourModel"
```

### Tabular Editor
1. Open Tabular Editor
2. Connect to Power BI Desktop (File > Open > From DB)
3. Run script: File > Run Script > Select ExtractSchema.csx

## Prerequisites

- Power BI Desktop (running with a model open)
- XMLA endpoint enabled (enhanced metadata format)
- For PowerShell: Microsoft.AnalysisServices.Tabular.dll
- For Python: pythonnet package

## Finding the XMLA Port

The XMLA endpoint typically uses ports in the range 50000-60000. Use one of these methods:

**Method 1: DAX Studio** (Easiest)
- Download DAX Studio (free)
- It automatically detects running PBI Desktop instances

**Method 2: PowerShell**
```powershell
Get-NetTCPConnection | Where-Object {$_.State -eq "Listen" -and $_.OwningProcess -in (Get-Process msmdsrv -ErrorAction SilentlyContinue).Id} | Select-Object -ExpandProperty LocalPort
```

**Method 3: External Tools**
- Tabular Editor registered as external tool will auto-connect

## Output Files

The scripts generate:
- **model-schema.json / model_*.bim** - Complete model definition (Model.bim format)
- **schema-report.txt** - Human-readable report with all details
- **measures.dax** - All DAX measures extracted
- **calculated_columns.dax** - All calculated columns (if any)
- **calculated_tables.dax** - All calculated tables (if any)
- **relationships.csv** - Relationship definitions
- **metadata.json** - Structured metadata for parsing

## Use Cases

- **Version Control** - Track model changes in Git
- **Documentation** - Auto-generate model documentation
- **CI/CD** - Integrate into deployment pipelines
- **Code Review** - Review DAX measures and calculations
- **Model Comparison** - Compare different versions
- **Migration** - Export for Azure Analysis Services or Power BI Service

## Troubleshooting

### Cannot find port
- Ensure Power BI Desktop is running with a model open
- Use DAX Studio to auto-detect the port
- Check Task Manager for msmdsrv.exe process

### Access denied
- Run PowerShell as Administrator
- Check Windows Firewall
- Verify antivirus isn't blocking

### Assembly not found
- Install Power BI Desktop (includes required DLLs)
- Or install SQL Server Management Studio
- Or install AMO/TOM via NuGet

## Learn More

See [tabular-editor-skill.md](tabular-editor-skill.md) for complete documentation.

---

**Created by:** @404-MikeNotFound
**Repository:** [404-MikeNotFound/404-MikeNotFound](https://github.com/404-MikeNotFound/404-MikeNotFound)
