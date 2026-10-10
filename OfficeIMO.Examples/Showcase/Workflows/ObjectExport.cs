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
        public DateTime PurchaseDate { get; init; }
        public decimal PurchaseCost { get; init; }
        public string Status { get; init; } = "";
    }

    internal static void Create(string folder) {
        Asset[] assets = {
            new() { AssetTag = "NW-1042", Category = "Laptop", AssignedTeam = "Delivery", PurchaseDate = new DateTime(2026, 1, 12), PurchaseCost = 1240, Status = "In use" },
            new() { AssetTag = "NW-1043", Category = "Monitor", AssignedTeam = "Support", PurchaseDate = new DateTime(2026, 1, 12), PurchaseCost = 320, Status = "In use" },
            new() { AssetTag = "NW-1044", Category = "Dock", AssignedTeam = "Engineering", PurchaseDate = new DateTime(2026, 2, 3), PurchaseCost = 185, Status = "In stock" },
            new() { AssetTag = "NW-1045", Category = "Laptop", AssignedTeam = "Operations", PurchaseDate = new DateTime(2026, 2, 17), PurchaseCost = 1390, Status = "In use" },
            new() { AssetTag = "NW-1046", Category = "Tablet", AssignedTeam = "Delivery", PurchaseDate = new DateTime(2026, 3, 2), PurchaseCost = 680, Status = "In repair" },
            new() { AssetTag = "NW-1047", Category = "Laptop", AssignedTeam = "Support", PurchaseDate = new DateTime(2026, 3, 9), PurchaseCost = 1240, Status = "In use" },
            new() { AssetTag = "NW-1048", Category = "Monitor", AssignedTeam = "Operations", PurchaseDate = new DateTime(2026, 3, 23), PurchaseCost = 320, Status = "In stock" },
            new() { AssetTag = "NW-1049", Category = "Phone", AssignedTeam = "Delivery", PurchaseDate = new DateTime(2026, 4, 6), PurchaseCost = 540, Status = "In use" },
            new() { AssetTag = "NW-1050", Category = "Dock", AssignedTeam = "Support", PurchaseDate = new DateTime(2026, 4, 6), PurchaseCost = 185, Status = "In use" },
            new() { AssetTag = "NW-1051", Category = "Laptop", AssignedTeam = "Engineering", PurchaseDate = new DateTime(2026, 4, 20), PurchaseCost = 1590, Status = "In stock" }
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
        for (int column = 1; column <= 6; column++) sheet.SetColumnWidth(column, 18);
        for (int row = 2; row <= assets.Length + 1; row++) {
            sheet.FormatCell(row, 4, "d mmm yyyy");
            sheet.FormatCell(row, 5, "#,##0.00");
        }
        sheet.Range("A1:F11").ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}