using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Exports typed application records as a filterable Excel table.</summary>
internal static class ObjectExport {
    private sealed class Asset {
        public string AssetTag { get; init; } = "";
        public string Category { get; init; } = "";
        public string AssignedTeam { get; init; } = "";
        public decimal PurchaseCost { get; init; }
    }

    internal static void Create(string folder) {
        Asset[] assets = {
            new() { AssetTag = "NW-1042", Category = "Laptop", AssignedTeam = "Delivery", PurchaseCost = 1240 },
            new() { AssetTag = "NW-1043", Category = "Monitor", AssignedTeam = "Support", PurchaseCost = 320 },
            new() { AssetTag = "NW-1044", Category = "Dock", AssignedTeam = "Engineering", PurchaseCost = 185 },
            new() { AssetTag = "NW-1045", Category = "Laptop", AssignedTeam = "Operations", PurchaseCost = 1390 },
            new() { AssetTag = "NW-1046", Category = "Tablet", AssignedTeam = "Delivery", PurchaseCost = 680 }
        };
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        document.AsFluent()
            .Sheet("Assets", sheet => sheet
                .RowsFrom(assets)
                .Table("Assets", table => table.Style(ExcelTableStyle.TableStyleMedium2))
                .AutoFit(columns: true, rows: false))
            .End();
        ExcelSheet sheet = document.Sheets[0];
        sheet.Freeze(topRows: 1);
        for (int column = 1; column <= 4; column++) sheet.SetColumnWidth(column, 24);
        for (int row = 2; row <= assets.Length + 1; row++) sheet.FormatCell(row, 4, "#,##0.00");
        sheet.Range("A1:D8").ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
