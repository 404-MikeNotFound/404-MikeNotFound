// Tabular Editor C# Script: Extract Power BI Schema
// Usage: Run this script from Tabular Editor while connected to a Power BI model
// File > Run Script > Select this file
// Or via command line: TabularEditor.exe "localhost:PORT" "DatabaseName" -S "ExtractSchema.csx"

using System;
using System.IO;
using System.Text;
using System.Linq;

// Configuration
var outputDirectory = @"C:\Temp\PBI_Schema_Export";
var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

// Create output directory
if (!Directory.Exists(outputDirectory)) {
    Directory.CreateDirectory(outputDirectory);
}

Info($"Starting schema extraction to: {outputDirectory}");

// 1. Export complete model schema to JSON (Model.bim format)
Info("Exporting model.bim (complete schema)...");
var json = Model.ToJson();
var jsonPath = Path.Combine(outputDirectory, $"model_{timestamp}.bim");
File.WriteAllText(jsonPath, json);
Info($"✓ Exported to: {jsonPath}");

// 2. Generate comprehensive schema report
Info("Generating schema report...");
var report = new StringBuilder();

report.AppendLine("═════════════════════════════════════════════════════════════════════════════");
report.AppendLine("POWER BI MODEL SCHEMA REPORT");
report.AppendLine("═════════════════════════════════════════════════════════════════════════════");
report.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
report.AppendLine($"Model: {Model.Database.Name}");
report.AppendLine($"Compatibility Level: {Model.Database.CompatibilityLevel}");
report.AppendLine();

// Summary
report.AppendLine("═════════════════════════════════════════════════════════════════════════════");
report.AppendLine("SUMMARY");
report.AppendLine("═════════════════════════════════════════════════════════════════════════════");
report.AppendLine($"Tables: {Model.Tables.Count}");
report.AppendLine($"Relationships: {Model.Relationships.Count}");
report.AppendLine($"Measures: {Model.AllMeasures.Count()}");
report.AppendLine($"Calculated Columns: {Model.AllColumns.Count(c => c.Type == ColumnType.Calculated)}");
report.AppendLine($"Calculated Tables: {Model.Tables.Count(t => !string.IsNullOrEmpty(t.Expression))}");
report.AppendLine($"Roles: {Model.Roles.Count}");
report.AppendLine($"Perspectives: {Model.Perspectives.Count}");
report.AppendLine();

// Tables
report.AppendLine("═════════════════════════════════════════════════════════════════════════════");
report.AppendLine("TABLES");
report.AppendLine("═════════════════════════════════════════════════════════════════════════════");

foreach(var table in Model.Tables.OrderBy(t => t.Name)) {
    report.AppendLine();
    report.AppendLine($"Table: {table.Name}");
    report.AppendLine("─────────────────────────────────────────────────────────────────────────────");
    report.AppendLine($"  Hidden: {table.IsHidden}");
    report.AppendLine($"  Type: {(string.IsNullOrEmpty(table.Expression) ? "Regular" : "Calculated")}");
    report.AppendLine($"  Columns: {table.Columns.Count}");
    report.AppendLine($"  Measures: {table.Measures.Count}");
    report.AppendLine($"  Hierarchies: {table.Hierarchies.Count}");
    report.AppendLine($"  Partitions: {table.Partitions.Count}");

    if (!string.IsNullOrEmpty(table.Description)) {
        report.AppendLine($"  Description: {table.Description}");
    }

    if (!string.IsNullOrEmpty(table.Expression)) {
        report.AppendLine($"  DAX Expression: {table.Expression}");
    }

    // Columns
    if (table.Columns.Count > 0) {
        report.AppendLine();
        report.AppendLine("  Columns:");
        foreach(var column in table.Columns.OrderBy(c => c.Name)) {
            var hidden = column.IsHidden ? " [HIDDEN]" : "";
            var key = column.IsKey ? " [KEY]" : "";
            var calculated = column.Type == ColumnType.Calculated ? " [CALCULATED]" : "";

            report.AppendLine($"    • {column.Name}: {column.DataType}{hidden}{key}{calculated}");

            if (!string.IsNullOrEmpty(column.FormatString)) {
                report.AppendLine($"      Format: {column.FormatString}");
            }

            if (column.Type == ColumnType.Calculated && !string.IsNullOrEmpty(column.Expression)) {
                report.AppendLine($"      Expression: {column.Expression}");
            }

            if (!string.IsNullOrEmpty(column.Description)) {
                report.AppendLine($"      Description: {column.Description}");
            }
        }
    }

    // Measures
    if (table.Measures.Count > 0) {
        report.AppendLine();
        report.AppendLine("  Measures:");
        foreach(var measure in table.Measures.OrderBy(m => m.Name)) {
            var hidden = measure.IsHidden ? " [HIDDEN]" : "";
            report.AppendLine($"    • [{measure.Name}]{hidden}");
            report.AppendLine($"      Expression: {measure.Expression}");

            if (!string.IsNullOrEmpty(measure.FormatString)) {
                report.AppendLine($"      Format: {measure.FormatString}");
            }

            if (!string.IsNullOrEmpty(measure.Description)) {
                report.AppendLine($"      Description: {measure.Description}");
            }

            if (!string.IsNullOrEmpty(measure.DisplayFolder)) {
                report.AppendLine($"      Display Folder: {measure.DisplayFolder}");
            }
        }
    }

    // Hierarchies
    if (table.Hierarchies.Count > 0) {
        report.AppendLine();
        report.AppendLine("  Hierarchies:");
        foreach(var hierarchy in table.Hierarchies.OrderBy(h => h.Name)) {
            var levels = string.Join(" → ", hierarchy.Levels.OrderBy(l => l.Ordinal).Select(l => l.Column.Name));
            report.AppendLine($"    • {hierarchy.Name}: {levels}");
        }
    }
}

// Relationships
report.AppendLine();
report.AppendLine("═════════════════════════════════════════════════════════════════════════════");
report.AppendLine("RELATIONSHIPS");
report.AppendLine("═════════════════════════════════════════════════════════════════════════════");

if (Model.Relationships.Count == 0) {
    report.AppendLine("No relationships defined");
} else {
    foreach(var rel in Model.Relationships.OrderBy(r => r.FromTable.Name)) {
        report.AppendLine();
        report.AppendLine($"Relationship: {rel.Name}");
        report.AppendLine("─────────────────────────────────────────────────────────────────────────────");
        report.AppendLine($"  From: {rel.FromTable.Name}[{rel.FromColumn.Name}]");
        report.AppendLine($"  To: {rel.ToTable.Name}[{rel.ToColumn.Name}]");
        report.AppendLine($"  Cardinality: {rel.FromCardinality} to {rel.ToCardinality}");
        report.AppendLine($"  Cross Filter: {rel.CrossFilteringBehavior}");
        report.AppendLine($"  Active: {rel.IsActive}");
        report.AppendLine($"  Security Filtering: {rel.SecurityFilteringBehavior}");
    }
}

// Roles (Row-Level Security)
if (Model.Roles.Count > 0) {
    report.AppendLine();
    report.AppendLine("═════════════════════════════════════════════════════════════════════════════");
    report.AppendLine("ROLES (Row-Level Security)");
    report.AppendLine("═════════════════════════════════════════════════════════════════════════════");

    foreach(var role in Model.Roles.OrderBy(r => r.Name)) {
        report.AppendLine();
        report.AppendLine($"Role: {role.Name}");
        report.AppendLine("─────────────────────────────────────────────────────────────────────────────");

        if (!string.IsNullOrEmpty(role.Description)) {
            report.AppendLine($"  Description: {role.Description}");
        }

        if (role.TablePermissions.Count > 0) {
            report.AppendLine("  Table Permissions:");
            foreach(var tp in role.TablePermissions) {
                report.AppendLine($"    Table: {tp.Table.Name}");
                if (!string.IsNullOrEmpty(tp.FilterExpression)) {
                    report.AppendLine($"    Filter: {tp.FilterExpression}");
                }
            }
        }
    }
}

// Perspectives
if (Model.Perspectives.Count > 0) {
    report.AppendLine();
    report.AppendLine("═════════════════════════════════════════════════════════════════════════════");
    report.AppendLine("PERSPECTIVES");
    report.AppendLine("═════════════════════════════════════════════════════════════════════════════");

    foreach(var perspective in Model.Perspectives.OrderBy(p => p.Name)) {
        report.AppendLine();
        report.AppendLine($"Perspective: {perspective.Name}");
        if (!string.IsNullOrEmpty(perspective.Description)) {
            report.AppendLine($"  Description: {perspective.Description}");
        }
    }
}

// Data Sources
if (Model.DataSources.Count > 0) {
    report.AppendLine();
    report.AppendLine("═════════════════════════════════════════════════════════════════════════════");
    report.AppendLine("DATA SOURCES");
    report.AppendLine("═════════════════════════════════════════════════════════════════════════════");

    foreach(var ds in Model.DataSources) {
        report.AppendLine();
        report.AppendLine($"Data Source: {ds.Name}");
        report.AppendLine($"  Type: {ds.Type}");
    }
}

// Save report
var reportPath = Path.Combine(outputDirectory, $"schema_report_{timestamp}.txt");
File.WriteAllText(reportPath, report.ToString());
Info($"✓ Report exported to: {reportPath}");

// 3. Export all DAX measures to separate file
Info("Exporting DAX measures...");
var measures = new StringBuilder();
measures.AppendLine("// ═════════════════════════════════════════════════════════════════════════════");
measures.AppendLine("// POWER BI MEASURES EXPORT");
measures.AppendLine($"// Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
measures.AppendLine($"// Model: {Model.Database.Name}");
measures.AppendLine("// ═════════════════════════════════════════════════════════════════════════════");
measures.AppendLine();

foreach(var table in Model.Tables.Where(t => t.Measures.Count > 0).OrderBy(t => t.Name)) {
    measures.AppendLine();
    measures.AppendLine($"// ─────────────────────────────────────────────────────────────────────────────");
    measures.AppendLine($"// Table: {table.Name}");
    measures.AppendLine($"// ─────────────────────────────────────────────────────────────────────────────");
    measures.AppendLine();

    foreach(var measure in table.Measures.OrderBy(m => m.Name)) {
        measures.AppendLine($"// Measure: [{measure.Name}]");
        if (!string.IsNullOrEmpty(measure.Description)) {
            measures.AppendLine($"// Description: {measure.Description}");
        }
        if (!string.IsNullOrEmpty(measure.FormatString)) {
            measures.AppendLine($"// Format: {measure.FormatString}");
        }
        if (!string.IsNullOrEmpty(measure.DisplayFolder)) {
            measures.AppendLine($"// Display Folder: {measure.DisplayFolder}");
        }

        measures.AppendLine($"[{measure.Name}] = ");
        measures.AppendLine(measure.Expression);
        measures.AppendLine();
    }
}

var measuresPath = Path.Combine(outputDirectory, $"measures_{timestamp}.dax");
File.WriteAllText(measuresPath, measures.ToString());
Info($"✓ Measures exported to: {measuresPath}");

// 4. Export calculated columns
Info("Exporting calculated columns...");
var calcColumns = Model.AllColumns.Where(c => c.Type == ColumnType.Calculated).ToList();

if (calcColumns.Any()) {
    var columns = new StringBuilder();
    columns.AppendLine("// ═════════════════════════════════════════════════════════════════════════════");
    columns.AppendLine("// CALCULATED COLUMNS EXPORT");
    columns.AppendLine($"// Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
    columns.AppendLine("// ═════════════════════════════════════════════════════════════════════════════");
    columns.AppendLine();

    foreach(var column in calcColumns.OrderBy(c => c.Table.Name).ThenBy(c => c.Name)) {
        columns.AppendLine($"// Table: {column.Table.Name}");
        columns.AppendLine($"// Column: {column.Name}");
        if (!string.IsNullOrEmpty(column.Description)) {
            columns.AppendLine($"// Description: {column.Description}");
        }
        columns.AppendLine($"{column.Name} = ");
        columns.AppendLine(column.Expression);
        columns.AppendLine();
    }

    var columnsPath = Path.Combine(outputDirectory, $"calculated_columns_{timestamp}.dax");
    File.WriteAllText(columnsPath, columns.ToString());
    Info($"✓ Calculated columns exported to: {columnsPath}");
}

// 5. Export calculated tables
Info("Exporting calculated tables...");
var calcTables = Model.Tables.Where(t => !string.IsNullOrEmpty(t.Expression)).ToList();

if (calcTables.Any()) {
    var tables = new StringBuilder();
    tables.AppendLine("// ═════════════════════════════════════════════════════════════════════════════");
    tables.AppendLine("// CALCULATED TABLES EXPORT");
    tables.AppendLine($"// Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
    tables.AppendLine("// ═════════════════════════════════════════════════════════════════════════════");
    tables.AppendLine();

    foreach(var table in calcTables.OrderBy(t => t.Name)) {
        tables.AppendLine($"// Table: {table.Name}");
        if (!string.IsNullOrEmpty(table.Description)) {
            tables.AppendLine($"// Description: {table.Description}");
        }
        tables.AppendLine($"{table.Name} = ");
        tables.AppendLine(table.Expression);
        tables.AppendLine();
    }

    var tablesPath = Path.Combine(outputDirectory, $"calculated_tables_{timestamp}.dax");
    File.WriteAllText(tablesPath, tables.ToString());
    Info($"✓ Calculated tables exported to: {tablesPath}");
}

// 6. Export relationships to CSV format
Info("Exporting relationships to CSV...");
var csv = new StringBuilder();
csv.AppendLine("FromTable,FromColumn,ToTable,ToColumn,FromCardinality,ToCardinality,CrossFilter,Active,SecurityFiltering");

foreach(var rel in Model.Relationships) {
    csv.AppendLine($"{rel.FromTable.Name},{rel.FromColumn.Name},{rel.ToTable.Name},{rel.ToColumn.Name}," +
                   $"{rel.FromCardinality},{rel.ToCardinality},{rel.CrossFilteringBehavior}," +
                   $"{rel.IsActive},{rel.SecurityFilteringBehavior}");
}

var csvPath = Path.Combine(outputDirectory, $"relationships_{timestamp}.csv");
File.WriteAllText(csvPath, csv.ToString());
Info($"✓ Relationships exported to: {csvPath}");

// Summary
Info("");
Info("═════════════════════════════════════════════════════════════════════════════");
Info("SCHEMA EXTRACTION COMPLETED SUCCESSFULLY!");
Info("═════════════════════════════════════════════════════════════════════════════");
Info($"Output directory: {outputDirectory}");
Info("");
Info("Files created:");
Info($"  • model_{timestamp}.bim");
Info($"  • schema_report_{timestamp}.txt");
Info($"  • measures_{timestamp}.dax");
if (calcColumns.Any()) Info($"  • calculated_columns_{timestamp}.dax");
if (calcTables.Any()) Info($"  • calculated_tables_{timestamp}.dax");
Info($"  • relationships_{timestamp}.csv");
