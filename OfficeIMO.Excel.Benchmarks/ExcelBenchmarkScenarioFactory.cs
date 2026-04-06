using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelBenchmarkScenarioFactory {
    private static readonly string[] Regions = ["North", "South", "East", "West", "Central"];
    private static readonly string[] Owners = ["Ava", "Noah", "Mia", "Liam", "Zoe", "Ethan", "Ivy", "Mason"];

    internal sealed class SalesRecord {
        public int Id { get; init; }
        public string Region { get; init; } = string.Empty;
        public string Owner { get; init; } = string.Empty;
        public DateTime CreatedOn { get; init; }
        public double Amount { get; init; }
        public int Units { get; init; }
        public bool Active { get; init; }
        public string Notes { get; init; } = string.Empty;
    }

    public static IReadOnlyList<SalesRecord> CreateSalesRecords(int count) {
        var records = new List<SalesRecord>(count);
        var start = new DateTime(2024, 1, 1, 8, 0, 0, DateTimeKind.Unspecified);

        for (int i = 0; i < count; i++) {
            records.Add(new SalesRecord {
                Id = i + 1,
                Region = Regions[i % Regions.Length],
                Owner = Owners[i % Owners.Length],
                CreatedOn = start.AddDays(i % 365).AddMinutes(i % 180),
                Amount = Math.Round(150 + ((i * 17.35) % 4500), 2),
                Units = 1 + (i % 24),
                Active = i % 3 != 0,
                Notes = $"Batch {(i % 12) + 1} / segment {(i % 7) + 1}"
            });
        }

        return records;
    }

    public static byte[] CreateWorkbookBytes(IReadOnlyList<SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            PopulateOfficeImoWorksheet(sheet, rows);
        }

        return stream.ToArray();
    }

    public static string BuildDataRange(int rowCount) => $"A1:H{rowCount + 1}";

    public static void PopulateOfficeImoWorksheet(ExcelSheet sheet, IReadOnlyList<SalesRecord> rows) {
        InsertOfficeImoObjects(sheet, rows);
        AddOfficeImoTable(sheet, rows.Count);
        AutoFitOfficeImoColumns(sheet);
    }

    public static void PopulateClosedXmlWorksheet(IXLWorksheet worksheet, IReadOnlyList<SalesRecord> rows) {
        var table = InsertClosedXmlTable(worksheet, rows);
        StyleClosedXmlTable(table);
        AutoFitClosedXmlColumns(worksheet);
    }

    public static void InsertOfficeImoObjects(ExcelSheet sheet, IReadOnlyList<SalesRecord> rows) {
        sheet.InsertObjects(rows,
            ("Id", item => item.Id),
            ("Region", item => item.Region),
            ("Owner", item => item.Owner),
            ("CreatedOn", item => item.CreatedOn),
            ("Amount", item => item.Amount),
            ("Units", item => item.Units),
            ("Active", item => item.Active),
            ("Notes", item => item.Notes));
    }

    public static void AddOfficeImoTable(ExcelSheet sheet, int rowCount) {
        sheet.AddTable(BuildDataRange(rowCount), hasHeader: true, name: "SalesData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
    }

    public static void AutoFitOfficeImoColumns(ExcelSheet sheet) {
        sheet.AutoFitColumns();
    }

    public static IXLTable InsertClosedXmlTable(IXLWorksheet worksheet, IReadOnlyList<SalesRecord> rows) {
        return worksheet.Cell(1, 1).InsertTable(rows);
    }

    public static void StyleClosedXmlTable(IXLTable table) {
        table.Theme = XLTableTheme.TableStyleMedium2;
    }

    public static void AutoFitClosedXmlColumns(IXLWorksheet worksheet) {
        worksheet.ColumnsUsed().AdjustToContents();
    }
}
