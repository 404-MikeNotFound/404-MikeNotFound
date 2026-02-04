#!/usr/bin/env python3
"""
Extract Power BI Desktop Schema via XMLA Endpoint

This script connects to Power BI Desktop through the XMLA endpoint and extracts
the complete tabular model schema including tables, columns, measures, relationships, and more.

Requirements:
    - pythonnet (install: pip install pythonnet)
    - Power BI Desktop running with a model open
    - Microsoft.AnalysisServices.Tabular.dll

Usage:
    python extract_pbi_schema.py --server localhost:PORT --database "ModelName"

Author: @404-MikeNotFound
"""

import argparse
import json
import sys
from pathlib import Path
from datetime import datetime

try:
    import clr
except ImportError:
    print("ERROR: pythonnet not installed")
    print("Install with: pip install pythonnet")
    sys.exit(1)


def find_tom_assembly():
    """Find the TOM (Tabular Object Model) assembly"""
    possible_paths = [
        r"C:\Program Files\Microsoft Power BI Desktop\bin\Microsoft.AnalysisServices.Tabular.dll",
        r"C:\Program Files (x86)\Microsoft Power BI Desktop\bin\Microsoft.AnalysisServices.Tabular.dll",
        r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.AnalysisServices.Tabular",
    ]

    for path_pattern in possible_paths:
        path = Path(path_pattern)
        if path.is_file():
            return str(path)
        elif path.is_dir():
            # Search in subdirectories
            dll_files = list(path.rglob("Microsoft.AnalysisServices.Tabular.dll"))
            if dll_files:
                return str(dll_files[0])

    return None


def load_tom():
    """Load the Tabular Object Model assembly"""
    tom_path = find_tom_assembly()

    if not tom_path:
        print("ERROR: Could not find Microsoft.AnalysisServices.Tabular.dll")
        print("\nPlease ensure one of the following is installed:")
        print("  - Power BI Desktop")
        print("  - SQL Server Management Studio")
        print("  - AMO/TOM NuGet package")
        sys.exit(1)

    print(f"Found TOM assembly at: {tom_path}")

    try:
        clr.AddReference(tom_path)
        from Microsoft.AnalysisServices.Tabular import Server, JsonSerializer
        return Server, JsonSerializer
    except Exception as e:
        print(f"ERROR loading TOM assembly: {e}")
        sys.exit(1)


def extract_schema(server_name, database_name=None, output_dir="./output"):
    """Extract Power BI model schema via XMLA endpoint"""

    Server, JsonSerializer = load_tom()

    # Create output directory
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True, parents=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    try:
        # Connect to server
        print(f"\nConnecting to server: {server_name}")
        server = Server()
        server.Connect(server_name)
        print("✓ Connected successfully!")

        # List available databases if not specified
        if not database_name:
            print("\nAvailable databases:")
            for db in server.Databases:
                print(f"  - {db.Name}")

            if server.Databases.Count == 1:
                database_name = server.Databases[0].Name
                print(f"\nAuto-selecting database: {database_name}")
            else:
                print("\nPlease specify --database parameter")
                server.Disconnect()
                sys.exit(1)

        # Get database
        database = server.Databases.FindByName(database_name)

        if database is None:
            print(f"ERROR: Database '{database_name}' not found")
            server.Disconnect()
            sys.exit(1)

        print(f"✓ Database found: {database_name}")

        model = database.Model

        # 1. Export full schema to JSON (Model.bim format)
        print("\nExtracting schema to JSON...")
        json_schema = JsonSerializer.SerializeDatabase(database)
        json_path = output_path / f"model_{timestamp}.json"

        with open(json_path, 'w', encoding='utf-8') as f:
            # Pretty print the JSON
            schema_dict = json.loads(json_schema)
            json.dump(schema_dict, f, indent=2, ensure_ascii=False)

        print(f"✓ Schema exported to: {json_path}")

        # 2. Generate detailed schema report
        print("\nGenerating schema report...")
        report = []

        report.append("=" * 80)
        report.append("POWER BI MODEL SCHEMA REPORT")
        report.append("=" * 80)
        report.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append(f"Server: {server_name}")
        report.append(f"Database: {database_name}")
        report.append(f"Compatibility Level: {database.CompatibilityLevel}")
        report.append("")

        # Summary
        report.append("=" * 80)
        report.append("SUMMARY")
        report.append("=" * 80)
        report.append(f"Tables: {model.Tables.Count}")
        report.append(f"Relationships: {model.Relationships.Count}")

        total_measures = sum(table.Measures.Count for table in model.Tables)
        total_columns = sum(table.Columns.Count for table in model.Tables)

        report.append(f"Measures: {total_measures}")
        report.append(f"Columns: {total_columns}")
        report.append(f"Roles: {model.Roles.Count}")
        report.append("")

        # Tables
        report.append("=" * 80)
        report.append("TABLES")
        report.append("=" * 80)

        for table in sorted(model.Tables, key=lambda t: t.Name):
            report.append("")
            report.append(f"Table: {table.Name}")
            report.append("-" * 40)
            report.append(f"  Hidden: {table.IsHidden}")
            report.append(f"  Columns: {table.Columns.Count}")
            report.append(f"  Measures: {table.Measures.Count}")
            report.append(f"  Hierarchies: {table.Hierarchies.Count}")
            report.append(f"  Partitions: {table.Partitions.Count}")

            if table.Description:
                report.append(f"  Description: {table.Description}")

            # Columns
            if table.Columns.Count > 0:
                report.append("")
                report.append("  Columns:")
                for column in sorted(table.Columns, key=lambda c: c.Name):
                    hidden = " [HIDDEN]" if column.IsHidden else ""
                    key = " [KEY]" if column.IsKey else ""
                    report.append(f"    • {column.Name}: {column.DataType}{hidden}{key}")

                    if column.FormatString:
                        report.append(f"      Format: {column.FormatString}")

            # Measures
            if table.Measures.Count > 0:
                report.append("")
                report.append("  Measures:")
                for measure in sorted(table.Measures, key=lambda m: m.Name):
                    hidden = " [HIDDEN]" if measure.IsHidden else ""
                    report.append(f"    • [{measure.Name}]{hidden}")
                    report.append(f"      Expression: {measure.Expression}")

                    if measure.FormatString:
                        report.append(f"      Format: {measure.FormatString}")

                    if measure.Description:
                        report.append(f"      Description: {measure.Description}")

            # Hierarchies
            if table.Hierarchies.Count > 0:
                report.append("")
                report.append("  Hierarchies:")
                for hierarchy in table.Hierarchies:
                    levels = " → ".join([level.Column.Name for level in hierarchy.Levels])
                    report.append(f"    • {hierarchy.Name}: {levels}")

        # Relationships
        report.append("")
        report.append("=" * 80)
        report.append("RELATIONSHIPS")
        report.append("=" * 80)

        if model.Relationships.Count == 0:
            report.append("No relationships defined")
        else:
            for rel in sorted(model.Relationships, key=lambda r: r.Name):
                report.append("")
                report.append(f"Relationship: {rel.Name}")
                report.append("-" * 40)
                report.append(f"  From: {rel.FromTable.Name}[{rel.FromColumn.Name}]")
                report.append(f"  To: {rel.ToTable.Name}[{rel.ToColumn.Name}]")
                report.append(f"  Cardinality: {rel.FromCardinality} to {rel.ToCardinality}")
                report.append(f"  Cross Filter: {rel.CrossFilteringBehavior}")
                report.append(f"  Active: {rel.IsActive}")
                report.append(f"  Security Filtering: {rel.SecurityFilteringBehavior}")

        # Roles
        if model.Roles.Count > 0:
            report.append("")
            report.append("=" * 80)
            report.append("ROLES (Row-Level Security)")
            report.append("=" * 80)

            for role in sorted(model.Roles, key=lambda r: r.Name):
                report.append("")
                report.append(f"Role: {role.Name}")
                report.append("-" * 40)

                if role.Description:
                    report.append(f"  Description: {role.Description}")

                if role.TablePermissions.Count > 0:
                    report.append("  Table Permissions:")
                    for tp in role.TablePermissions:
                        report.append(f"    Table: {tp.Table.Name}")
                        if tp.FilterExpression:
                            report.append(f"    Filter: {tp.FilterExpression}")

        # Save report
        report_path = output_path / f"schema_report_{timestamp}.txt"
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report))

        print(f"✓ Report exported to: {report_path}")

        # 3. Export measures to separate file
        print("\nExporting measures...")
        measures_report = []
        measures_report.append("// " + "=" * 77)
        measures_report.append("// POWER BI MEASURES EXPORT")
        measures_report.append(f"// Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        measures_report.append(f"// Model: {database_name}")
        measures_report.append("// " + "=" * 77)
        measures_report.append("")

        for table in sorted(model.Tables, key=lambda t: t.Name):
            if table.Measures.Count > 0:
                measures_report.append("")
                measures_report.append(f"// Table: {table.Name}")
                measures_report.append("")

                for measure in sorted(table.Measures, key=lambda m: m.Name):
                    measures_report.append(f"// Measure: [{measure.Name}]")
                    if measure.Description:
                        measures_report.append(f"// Description: {measure.Description}")
                    if measure.FormatString:
                        measures_report.append(f"// Format: {measure.FormatString}")

                    measures_report.append(f"[{measure.Name}] = ")
                    measures_report.append(measure.Expression)
                    measures_report.append("")

        measures_path = output_path / f"measures_{timestamp}.dax"
        with open(measures_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(measures_report))

        print(f"✓ Measures exported to: {measures_path}")

        # 4. Export metadata to JSON
        print("\nExporting metadata...")
        metadata = {
            "database": database_name,
            "compatibility_level": database.CompatibilityLevel,
            "generated": datetime.now().isoformat(),
            "tables": [],
            "relationships": [],
            "roles": []
        }

        for table in model.Tables:
            table_info = {
                "name": table.Name,
                "hidden": table.IsHidden,
                "columns": [
                    {
                        "name": col.Name,
                        "type": str(col.DataType),
                        "hidden": col.IsHidden
                    }
                    for col in table.Columns
                ],
                "measures": [
                    {
                        "name": m.Name,
                        "expression": m.Expression,
                        "format": m.FormatString if m.FormatString else None
                    }
                    for m in table.Measures
                ]
            }
            metadata["tables"].append(table_info)

        for rel in model.Relationships:
            rel_info = {
                "name": rel.Name,
                "from": f"{rel.FromTable.Name}[{rel.FromColumn.Name}]",
                "to": f"{rel.ToTable.Name}[{rel.ToColumn.Name}]",
                "cardinality": f"{rel.FromCardinality} to {rel.ToCardinality}",
                "active": rel.IsActive
            }
            metadata["relationships"].append(rel_info)

        for role in model.Roles:
            role_info = {
                "name": role.Name,
                "description": role.Description if role.Description else None
            }
            metadata["roles"].append(role_info)

        metadata_path = output_path / f"metadata_{timestamp}.json"
        with open(metadata_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2, ensure_ascii=False)

        print(f"✓ Metadata exported to: {metadata_path}")

        # Disconnect
        server.Disconnect()

        # Summary
        print("\n" + "=" * 80)
        print("SCHEMA EXTRACTION COMPLETED SUCCESSFULLY!")
        print("=" * 80)
        print(f"\nOutput directory: {output_path.absolute()}")
        print("\nFiles created:")
        print(f"  • model_{timestamp}.json         (Complete model definition)")
        print(f"  • schema_report_{timestamp}.txt   (Human-readable report)")
        print(f"  • measures_{timestamp}.dax        (All DAX measures)")
        print(f"  • metadata_{timestamp}.json       (Structured metadata)")

        return 0

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()

        if 'server' in locals() and server.Connected:
            server.Disconnect()

        return 1


def main():
    parser = argparse.ArgumentParser(
        description="Extract Power BI Desktop schema via XMLA endpoint",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python extract_pbi_schema.py --server localhost:12345
  python extract_pbi_schema.py --server localhost:12345 --database "Sales Model"
  python extract_pbi_schema.py --server localhost:12345 --output ./exports

Note: Power BI Desktop must be running with a model open.
        """
    )

    parser.add_argument(
        '--server', '-s',
        required=True,
        help='XMLA endpoint connection string (e.g., localhost:12345)'
    )

    parser.add_argument(
        '--database', '-d',
        help='Database/model name (auto-detected if only one is available)'
    )

    parser.add_argument(
        '--output', '-o',
        default='./output',
        help='Output directory (default: ./output)'
    )

    args = parser.parse_args()

    return extract_schema(args.server, args.database, args.output)


if __name__ == "__main__":
    sys.exit(main())
