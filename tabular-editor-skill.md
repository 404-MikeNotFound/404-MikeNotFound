# Tabular Editor Skill: Connect to Power BI Desktop via XMLA Endpoint and Extract Schema

## Overview
This guide demonstrates how to connect to Power BI Desktop through the XMLA (XML for Analysis) endpoint to extract the tabular model schema. This is useful for documentation, version control, and automated model analysis.

## Prerequisites

1. **Power BI Desktop** (Latest version)
2. **Tabular Editor 2** or **Tabular Editor 3** (Download from https://tabulareditor.com/)
3. **Power BI Desktop XMLA Endpoint** enabled
4. **AMO (Analysis Management Objects)** or **TOM (Tabular Object Model)** libraries for scripting

## Step 1: Enable XMLA Endpoint in Power BI Desktop

1. Open Power BI Desktop
2. Go to **File > Options and Settings > Options**
3. Navigate to **Preview Features**
4. Enable **"Store datasets using enhanced metadata format"**
5. Restart Power BI Desktop
6. Open your .pbix file

## Step 2: Find the XMLA Connection String

The XMLA endpoint for Power BI Desktop uses a local port. To find it:

1. Open your .pbix file in Power BI Desktop
2. The connection string format is:
   ```
   localhost:<port>
   ```

### Finding the Port Number

**Method 1: Using Power BI Desktop (External Tools)**
- The port is typically discovered automatically by external tools like Tabular Editor

**Method 2: Using PowerShell**
```powershell
# Get the Power BI Desktop process and find the XMLA port
Get-NetTCPConnection | Where-Object {$_.State -eq "Listen" -and $_.OwningProcess -in (Get-Process msmdsrv -ErrorAction SilentlyContinue).Id} | Select-Object -ExpandProperty LocalPort
```

**Method 3: Check Windows Registry or DAX Studio**
- Use DAX Studio (free tool) which automatically detects running PBI Desktop instances
- Or check: typically ports range from 50000-60000

## Step 3: Connect Using Tabular Editor

### GUI Method (Tabular Editor 2/3)

1. Open Tabular Editor
2. Go to **File > Open > From DB**
3. Enter connection details:
   - **Server:** `localhost:PORT` (e.g., `localhost:55555`)
   - **Database:** Select your model from dropdown
4. Click **Connect**

### Command Line Method (Tabular Editor 2)

```bash
TabularEditor.exe "localhost:PORT" "DatabaseName" -S "ExtractSchema.csx"
```

## Step 4: Extract Schema Using Scripts

### Option A: Tabular Editor C# Script (ExtractSchema.csx)

Create a file named `ExtractSchema.csx`:

```csharp
// Extract complete model schema to JSON (Model.bim format)
var json = Model.ToJson();
System.IO.File.WriteAllText(@"C:\Temp\ModelSchema.json", json);

// Extract individual components
var sb = new System.Text.StringBuilder();

// Tables
sb.AppendLine("=== TABLES ===");
foreach(var table in Model.Tables) {
    sb.AppendLine($"Table: {table.Name}");
    sb.AppendLine($"  Partitions: {table.Partitions.Count}");
    sb.AppendLine($"  Columns: {table.Columns.Count}");
    sb.AppendLine($"  Measures: {table.Measures.Count}");
    sb.AppendLine();
}

// Relationships
sb.AppendLine("=== RELATIONSHIPS ===");
foreach(var rel in Model.Relationships) {
    sb.AppendLine($"{rel.FromTable.Name}[{rel.FromColumn.Name}] -> {rel.ToTable.Name}[{rel.ToColumn.Name}]");
    sb.AppendLine($"  Cardinality: {rel.FromCardinality} to {rel.ToCardinality}");
    sb.AppendLine($"  Cross Filter: {rel.CrossFilteringBehavior}");
    sb.AppendLine();
}

// Measures
sb.AppendLine("=== MEASURES ===");
foreach(var table in Model.Tables) {
    foreach(var measure in table.Measures) {
        sb.AppendLine($"[{measure.Name}] in {table.Name}");
        sb.AppendLine($"  Expression: {measure.Expression}");
        sb.AppendLine($"  Format: {measure.FormatString}");
        sb.AppendLine();
    }
}

// Save to file
System.IO.File.WriteAllText(@"C:\Temp\SchemaReport.txt", sb.ToString());

Info("Schema extracted successfully!");
```

### Option B: PowerShell Script with AMO/TOM

Create `Extract-PBISchema.ps1`:

```powershell
# Load TOM assembly
Add-Type -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.AnalysisServices.Tabular\v4.0_15.0.0.0__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.dll"

# Connection parameters
$serverName = "localhost:PORT"  # Replace PORT with actual port number
$databaseName = "YourModelName"  # Replace with your model name

try {
    # Connect to server
    $server = New-Object Microsoft.AnalysisServices.Tabular.Server
    $server.Connect($serverName)

    Write-Host "Connected to: $serverName" -ForegroundColor Green

    # Get database
    $database = $server.Databases.FindByName($databaseName)

    if ($null -eq $database) {
        Write-Host "Database not found: $databaseName" -ForegroundColor Red
        $server.Disconnect()
        exit
    }

    # Get model
    $model = $database.Model

    # Extract schema to JSON
    $json = [Microsoft.AnalysisServices.Tabular.JsonSerializer]::SerializeDatabase($database)
    $outputPath = "C:\Temp\PBI_Schema_Export.json"
    $json | Out-File -FilePath $outputPath -Encoding UTF8

    Write-Host "Schema exported to: $outputPath" -ForegroundColor Green

    # Generate schema report
    $report = @()

    $report += "=== POWER BI MODEL SCHEMA REPORT ==="
    $report += "Server: $serverName"
    $report += "Database: $databaseName"
    $report += "Compatibility Level: $($database.CompatibilityLevel)"
    $report += ""

    $report += "=== TABLES ==="
    foreach ($table in $model.Tables) {
        $report += "Table: $($table.Name)"
        $report += "  Type: $($table.GetType().Name)"
        $report += "  Columns: $($table.Columns.Count)"
        $report += "  Measures: $($table.Measures.Count)"
        $report += "  Hierarchies: $($table.Hierarchies.Count)"
        $report += "  Partitions: $($table.Partitions.Count)"

        # List columns
        foreach ($column in $table.Columns) {
            $report += "    Column: $($column.Name) ($($column.DataType))"
        }

        # List measures
        foreach ($measure in $table.Measures) {
            $report += "    Measure: $($measure.Name)"
        }
        $report += ""
    }

    $report += "=== RELATIONSHIPS ==="
    foreach ($rel in $model.Relationships) {
        $report += "Relationship: $($rel.Name)"
        $report += "  From: $($rel.FromTable.Name)[$($rel.FromColumn.Name)]"
        $report += "  To: $($rel.ToTable.Name)[$($rel.ToColumn.Name)]"
        $report += "  Cardinality: $($rel.FromCardinality) to $($rel.ToCardinality)"
        $report += "  Cross Filter: $($rel.CrossFilteringBehavior)"
        $report += "  Active: $($rel.IsActive)"
        $report += ""
    }

    $report += "=== MEASURES ==="
    foreach ($table in $model.Tables) {
        foreach ($measure in $table.Measures) {
            $report += "[$($measure.Name)] in $($table.Name)"
            $report += "  Expression: $($measure.Expression)"
            $report += "  Format: $($measure.FormatString)"
            $report += ""
        }
    }

    # Save report
    $reportPath = "C:\Temp\PBI_Schema_Report.txt"
    $report | Out-File -FilePath $reportPath -Encoding UTF8

    Write-Host "Report exported to: $reportPath" -ForegroundColor Green

    # Disconnect
    $server.Disconnect()

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
```

### Option C: Python Script with pythonnet

Create `extract_pbi_schema.py`:

```python
import clr
import sys
import json
from pathlib import Path

# Add reference to TOM assembly
clr.AddReference("Microsoft.AnalysisServices.Tabular")
from Microsoft.AnalysisServices.Tabular import Server, JsonSerializer

def extract_schema(server_name, database_name, output_dir="./output"):
    """Extract Power BI model schema via XMLA endpoint"""

    Path(output_dir).mkdir(exist_ok=True)

    # Connect to server
    server = Server()
    server.Connect(server_name)
    print(f"Connected to: {server_name}")

    # Get database
    database = server.Databases.FindByName(database_name)
    if database is None:
        print(f"Database '{database_name}' not found")
        return

    model = database.Model

    # Export full schema to JSON
    json_schema = JsonSerializer.SerializeDatabase(database)
    output_file = Path(output_dir) / "schema.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(json_schema)
    print(f"Schema exported to: {output_file}")

    # Extract metadata
    metadata = {
        "database": database_name,
        "compatibility_level": database.CompatibilityLevel,
        "tables": [],
        "relationships": [],
        "measures": []
    }

    # Tables
    for table in model.Tables:
        table_info = {
            "name": table.Name,
            "columns": [{"name": col.Name, "type": str(col.DataType)} for col in table.Columns],
            "measures": [{"name": m.Name, "expression": m.Expression} for m in table.Measures],
            "partitions": table.Partitions.Count
        }
        metadata["tables"].append(table_info)

    # Relationships
    for rel in model.Relationships:
        rel_info = {
            "name": rel.Name,
            "from": f"{rel.FromTable.Name}[{rel.FromColumn.Name}]",
            "to": f"{rel.ToTable.Name}[{rel.ToColumn.Name}]",
            "cardinality": f"{rel.FromCardinality} to {rel.ToCardinality}",
            "active": rel.IsActive
        }
        metadata["relationships"].append(rel_info)

    # Save metadata
    metadata_file = Path(output_dir) / "metadata.json"
    with open(metadata_file, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, indent=2)
    print(f"Metadata exported to: {metadata_file}")

    server.Disconnect()

if __name__ == "__main__":
    # Example usage
    SERVER = "localhost:PORT"  # Replace with actual port
    DATABASE = "YourModelName"  # Replace with your model name

    extract_schema(SERVER, DATABASE)
```

## Step 5: Automated Extraction via External Tools Integration

Power BI Desktop supports External Tools. You can register Tabular Editor as an external tool:

1. Create a `.pbitool.json` file
2. Place it in: `C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools\`

Example `TabularEditor.pbitool.json`:
```json
{
  "version": "1.0.0",
  "name": "Tabular Editor",
  "description": "Open model in Tabular Editor",
  "path": "C:\\Program Files\\Tabular Editor\\TabularEditor.exe",
  "arguments": "\"%server%\" \"%database%\"",
  "iconData": "..."
}
```

## Output Format

The extracted schema will contain:

- **Tables**: Names, column definitions, data types
- **Columns**: Names, data types, descriptions, formatting
- **Measures**: DAX expressions, format strings, display folders
- **Relationships**: From/To tables and columns, cardinality, filter direction
- **Hierarchies**: Level definitions
- **Partitions**: M queries, source expressions
- **Roles**: RLS expressions
- **Perspectives**: Included objects
- **Translations**: Culture-specific names
- **Data Sources**: Connection information (if present)

## Use Cases

1. **Version Control**: Track model changes over time
2. **Documentation**: Generate automated documentation
3. **CI/CD**: Integrate into deployment pipelines
4. **Model Comparison**: Compare different versions
5. **Code Review**: Review DAX measures and calculated columns
6. **Migration**: Export schema for migration to Azure AS or Power BI Service

## Troubleshooting

### Cannot Connect to XMLA Endpoint
- Ensure Power BI Desktop is running with a model open
- Check that the port is correct
- Verify enhanced metadata format is enabled
- Try restarting Power BI Desktop

### Access Denied
- Run as Administrator if needed
- Check Windows Firewall settings
- Ensure no antivirus is blocking the connection

### Model Not Found
- Verify the database name (usually the .pbix filename without extension)
- List available databases using: `$server.Databases | Select-Object Name`

## References

- [Tabular Editor Documentation](https://docs.tabulareditor.com/)
- [Power BI XMLA Endpoint Documentation](https://docs.microsoft.com/en-us/power-bi/admin/service-premium-connect-tools)
- [Tabular Object Model (TOM) Reference](https://docs.microsoft.com/en-us/analysis-services/tom/introduction-to-the-tabular-object-model-tom-in-analysis-services-amo)

## License

This skill document is provided as-is for educational purposes.

---

**Created by:** @404-MikeNotFound
**Last Updated:** 2026-02-04
