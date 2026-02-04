<#
.SYNOPSIS
    Extract Power BI Desktop schema via XMLA endpoint

.DESCRIPTION
    This script connects to a running Power BI Desktop instance through the XMLA endpoint
    and extracts the complete tabular model schema including tables, relationships, measures, and more.

.PARAMETER ServerName
    The XMLA endpoint connection string (e.g., "localhost:12345")

.PARAMETER DatabaseName
    The name of the Power BI model/database (usually the .pbix filename without extension)

.PARAMETER OutputDirectory
    Directory where schema files will be saved (default: .\output)

.EXAMPLE
    .\Extract-PBISchema.ps1 -ServerName "localhost:54321" -DatabaseName "Sales Model"

.NOTES
    Author: @404-MikeNotFound
    Requires: Microsoft.AnalysisServices.Tabular assembly
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$ServerName,

    [Parameter(Mandatory=$false)]
    [string]$DatabaseName,

    [Parameter(Mandatory=$false)]
    [string]$OutputDirectory = ".\output"
)

# Function to find TOM assembly
function Find-TOMAssembly {
    $possiblePaths = @(
        # Power BI Desktop installations
        "C:\Program Files\Microsoft Power BI Desktop\bin\Microsoft.AnalysisServices.Tabular.dll",
        "C:\Program Files (x86)\Microsoft Power BI Desktop\bin\Microsoft.AnalysisServices.Tabular.dll",

        # Visual Studio installations
        "C:\Program Files\Microsoft SQL Server\*\Tools\Binn\ManagementStudio\Microsoft.AnalysisServices.Tabular.dll",

        # GAC
        "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.AnalysisServices.Tabular\*\Microsoft.AnalysisServices.Tabular.dll",

        # NuGet package locations
        "$env:USERPROFILE\.nuget\packages\microsoft.analysisservices.tabular\*\lib\*\Microsoft.AnalysisServices.Tabular.dll"
    )

    foreach ($pathPattern in $possiblePaths) {
        $resolved = Resolve-Path $pathPattern -ErrorAction SilentlyContinue
        if ($resolved) {
            return $resolved | Select-Object -First 1 -ExpandProperty Path
        }
    }

    return $null
}

# Load TOM assembly
Write-Host "Loading Tabular Object Model (TOM) assembly..." -ForegroundColor Cyan

$tomPath = Find-TOMAssembly

if (-not $tomPath) {
    Write-Host "ERROR: Could not find Microsoft.AnalysisServices.Tabular.dll" -ForegroundColor Red
    Write-Host "Please ensure one of the following is installed:" -ForegroundColor Yellow
    Write-Host "  - Power BI Desktop" -ForegroundColor Yellow
    Write-Host "  - SQL Server Management Studio" -ForegroundColor Yellow
    Write-Host "  - AMO/TOM NuGet package" -ForegroundColor Yellow
    exit 1
}

Write-Host "Found TOM assembly at: $tomPath" -ForegroundColor Green

try {
    Add-Type -Path $tomPath
} catch {
    Write-Host "ERROR loading TOM assembly: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Create output directory
if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

try {
    # Connect to server
    Write-Host "`nConnecting to server: $ServerName" -ForegroundColor Cyan
    $server = New-Object Microsoft.AnalysisServices.Tabular.Server
    $server.Connect($ServerName)
    Write-Host "Connected successfully!" -ForegroundColor Green

    # List available databases if DatabaseName not provided
    if (-not $DatabaseName) {
        Write-Host "`nAvailable databases:" -ForegroundColor Yellow
        foreach ($db in $server.Databases) {
            Write-Host "  - $($db.Name)" -ForegroundColor White
        }

        if ($server.Databases.Count -eq 1) {
            $DatabaseName = $server.Databases[0].Name
            Write-Host "`nAuto-selecting database: $DatabaseName" -ForegroundColor Cyan
        } else {
            $server.Disconnect()
            Write-Host "`nPlease specify -DatabaseName parameter" -ForegroundColor Red
            exit 1
        }
    }

    # Get database
    $database = $server.Databases.FindByName($DatabaseName)

    if ($null -eq $database) {
        Write-Host "ERROR: Database '$DatabaseName' not found" -ForegroundColor Red
        $server.Disconnect()
        exit 1
    }

    Write-Host "Database found: $DatabaseName" -ForegroundColor Green

    # Get model
    $model = $database.Model

    # Extract full schema to JSON (Model.bim format)
    Write-Host "`nExtracting schema to JSON..." -ForegroundColor Cyan
    $json = [Microsoft.AnalysisServices.Tabular.JsonSerializer]::SerializeDatabase($database)
    $jsonPath = Join-Path $OutputDirectory "model-schema.json"
    $json | Out-File -FilePath $jsonPath -Encoding UTF8
    Write-Host "Schema exported to: $jsonPath" -ForegroundColor Green

    # Generate detailed schema report
    Write-Host "`nGenerating schema report..." -ForegroundColor Cyan
    $report = @()

    $report += "═" * 80
    $report += "POWER BI MODEL SCHEMA REPORT"
    $report += "═" * 80
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "Server: $ServerName"
    $report += "Database: $DatabaseName"
    $report += "Compatibility Level: $($database.CompatibilityLevel)"
    $report += "Model Type: $(if ($model.DefaultMode -eq 'Import') { 'Import' } else { 'DirectQuery/Composite' })"
    $report += ""

    # Summary
    $report += "═" * 80
    $report += "SUMMARY"
    $report += "═" * 80
    $report += "Tables: $($model.Tables.Count)"
    $report += "Relationships: $($model.Relationships.Count)"
    $totalMeasures = ($model.Tables | ForEach-Object { $_.Measures.Count } | Measure-Object -Sum).Sum
    $report += "Measures: $totalMeasures"
    $totalColumns = ($model.Tables | ForEach-Object { $_.Columns.Count } | Measure-Object -Sum).Sum
    $report += "Columns: $totalColumns"
    $report += "Roles: $($model.Roles.Count)"
    $report += ""

    # Tables
    $report += "═" * 80
    $report += "TABLES"
    $report += "═" * 80
    foreach ($table in $model.Tables | Sort-Object Name) {
        $report += ""
        $report += "Table: $($table.Name)"
        $report += "─" * 40
        $report += "  Type: $(if ($table.Partitions[0].SourceType -eq 'M') { 'Power Query' } else { 'Other' })"
        $report += "  Hidden: $($table.IsHidden)"
        $report += "  Columns: $($table.Columns.Count)"
        $report += "  Measures: $($table.Measures.Count)"
        $report += "  Hierarchies: $($table.Hierarchies.Count)"
        $report += "  Partitions: $($table.Partitions.Count)"

        if ($table.Description) {
            $report += "  Description: $($table.Description)"
        }

        # Columns
        if ($table.Columns.Count -gt 0) {
            $report += ""
            $report += "  Columns:"
            foreach ($column in $table.Columns | Sort-Object Name) {
                $hidden = if ($column.IsHidden) { " [HIDDEN]" } else { "" }
                $key = if ($column.IsKey) { " [KEY]" } else { "" }
                $report += "    - $($column.Name): $($column.DataType)$hidden$key"
                if ($column.FormatString) {
                    $report += "      Format: $($column.FormatString)"
                }
            }
        }

        # Measures
        if ($table.Measures.Count -gt 0) {
            $report += ""
            $report += "  Measures:"
            foreach ($measure in $table.Measures | Sort-Object Name) {
                $hidden = if ($measure.IsHidden) { " [HIDDEN]" } else { "" }
                $report += "    - [$($measure.Name)]$hidden"
                $report += "      Expression: $($measure.Expression)"
                if ($measure.FormatString) {
                    $report += "      Format: $($measure.FormatString)"
                }
                if ($measure.Description) {
                    $report += "      Description: $($measure.Description)"
                }
            }
        }

        # Hierarchies
        if ($table.Hierarchies.Count -gt 0) {
            $report += ""
            $report += "  Hierarchies:"
            foreach ($hierarchy in $table.Hierarchies) {
                $report += "    - $($hierarchy.Name)"
                $levels = $hierarchy.Levels | ForEach-Object { $_.Column.Name }
                $report += "      Levels: $($levels -join ' > ')"
            }
        }
    }

    # Relationships
    $report += ""
    $report += "═" * 80
    $report += "RELATIONSHIPS"
    $report += "═" * 80
    if ($model.Relationships.Count -eq 0) {
        $report += "No relationships defined"
    } else {
        foreach ($rel in $model.Relationships | Sort-Object Name) {
            $report += ""
            $report += "Relationship: $($rel.Name)"
            $report += "─" * 40
            $report += "  From: $($rel.FromTable.Name)[$($rel.FromColumn.Name)]"
            $report += "  To: $($rel.ToTable.Name)[$($rel.ToColumn.Name)]"
            $report += "  Cardinality: $($rel.FromCardinality) to $($rel.ToCardinality)"
            $report += "  Cross Filter: $($rel.CrossFilteringBehavior)"
            $report += "  Active: $($rel.IsActive)"
            $report += "  Security: $($rel.SecurityFilteringBehavior)"
        }
    }

    # Roles (RLS)
    if ($model.Roles.Count -gt 0) {
        $report += ""
        $report += "═" * 80
        $report += "ROLES (Row-Level Security)"
        $report += "═" * 80
        foreach ($role in $model.Roles | Sort-Object Name) {
            $report += ""
            $report += "Role: $($role.Name)"
            $report += "─" * 40
            if ($role.Description) {
                $report += "  Description: $($role.Description)"
            }
            if ($role.TablePermissions.Count -gt 0) {
                $report += "  Table Permissions:"
                foreach ($tp in $role.TablePermissions) {
                    $report += "    Table: $($tp.Table.Name)"
                    if ($tp.FilterExpression) {
                        $report += "    Filter: $($tp.FilterExpression)"
                    }
                }
            }
        }
    }

    # Data Sources
    if ($model.DataSources.Count -gt 0) {
        $report += ""
        $report += "═" * 80
        $report += "DATA SOURCES"
        $report += "═" * 80
        foreach ($ds in $model.DataSources) {
            $report += ""
            $report += "Data Source: $($ds.Name)"
            $report += "  Type: $($ds.Type)"
            if ($ds.Type -eq 'Structured') {
                $report += "  Protocol: $($ds.Protocol)"
            }
        }
    }

    # Expressions (M queries)
    $report += ""
    $report += "═" * 80
    $report += "PARTITION SOURCES (Power Query M)"
    $report += "═" * 80
    foreach ($table in $model.Tables | Sort-Object Name) {
        foreach ($partition in $table.Partitions) {
            if ($partition.SourceType -eq 'M') {
                $report += ""
                $report += "Table: $($table.Name)"
                $report += "Partition: $($partition.Name)"
                $report += "─" * 40
                $report += $partition.Source.Expression
            }
        }
    }

    # Save report
    $reportPath = Join-Path $OutputDirectory "schema-report.txt"
    $report | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "Report exported to: $reportPath" -ForegroundColor Green

    # Export measures to separate file
    Write-Host "`nExporting measures..." -ForegroundColor Cyan
    $measuresReport = @()
    $measuresReport += "POWER BI MEASURES EXPORT"
    $measuresReport += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $measuresReport += ""

    foreach ($table in $model.Tables | Sort-Object Name) {
        if ($table.Measures.Count -gt 0) {
            $measuresReport += ""
            $measuresReport += "// Table: $($table.Name)"
            $measuresReport += ""
            foreach ($measure in $table.Measures | Sort-Object Name) {
                $measuresReport += "[$($measure.Name)] ="
                $measuresReport += $measure.Expression
                if ($measure.FormatString) {
                    $measuresReport += "// Format: $($measure.FormatString)"
                }
                $measuresReport += ""
            }
        }
    }

    $measuresPath = Join-Path $OutputDirectory "measures.dax"
    $measuresReport | Out-File -FilePath $measuresPath -Encoding UTF8
    Write-Host "Measures exported to: $measuresPath" -ForegroundColor Green

    # Export relationships to CSV
    Write-Host "`nExporting relationships to CSV..." -ForegroundColor Cyan
    $relationshipsData = @()
    foreach ($rel in $model.Relationships) {
        $relationshipsData += [PSCustomObject]@{
            Name = $rel.Name
            FromTable = $rel.FromTable.Name
            FromColumn = $rel.FromColumn.Name
            ToTable = $rel.ToTable.Name
            ToColumn = $rel.ToColumn.Name
            FromCardinality = $rel.FromCardinality
            ToCardinality = $rel.ToCardinality
            CrossFilter = $rel.CrossFilteringBehavior
            Active = $rel.IsActive
        }
    }

    if ($relationshipsData.Count -gt 0) {
        $relationshipsPath = Join-Path $OutputDirectory "relationships.csv"
        $relationshipsData | Export-Csv -Path $relationshipsPath -NoTypeInformation -Encoding UTF8
        Write-Host "Relationships exported to: $relationshipsPath" -ForegroundColor Green
    }

    # Disconnect
    $server.Disconnect()

    Write-Host "`n" + ("═" * 80) -ForegroundColor Green
    Write-Host "SCHEMA EXTRACTION COMPLETED SUCCESSFULLY" -ForegroundColor Green
    Write-Host ("═" * 80) -ForegroundColor Green
    Write-Host "`nOutput directory: $(Resolve-Path $OutputDirectory)" -ForegroundColor Cyan
    Write-Host "`nFiles created:" -ForegroundColor White
    Write-Host "  - model-schema.json     (Complete model definition)" -ForegroundColor White
    Write-Host "  - schema-report.txt     (Human-readable report)" -ForegroundColor White
    Write-Host "  - measures.dax          (All DAX measures)" -ForegroundColor White
    if ($relationshipsData.Count -gt 0) {
        Write-Host "  - relationships.csv     (Relationships data)" -ForegroundColor White
    }

} catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red

    if ($server -and $server.Connected) {
        $server.Disconnect()
    }

    exit 1
}
