using System.Data;
using OfficeIMO.Data;
using OfficeIMO.Excel;

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".xlsx");
try {
    using (var asyncRows = new MemoryStream()) {
        ExcelDataSetImportResult result = await ExcelDocument.WriteRowsAsync(
            asyncRows,
            CreateRowsAsync(),
            ["Region", "Revenue"],
            static (writer, row) => writer.Write(row.Region).Write(row.Revenue));
        if (result.Range != "A1:B3" || result.RowCount != 2) {
            throw new InvalidOperationException("The asynchronous row writer returned an unexpected range.");
        }
    }

    using (ExcelDocument document = ExcelDocument.Create(path)) {
        var sales = new DataTable("Sales");
        sales.Columns.Add("Region", typeof(string));
        sales.Columns.Add("Revenue", typeof(decimal));
        sales.Rows.Add("North", 1250000M);
        sales.Rows.Add("South", 980000M);

        ExcelSheet sheet = document.AddWorksheet("NativeAOT data");
        string range = sheet.InsertDataTableAsTable(sales, tableName: "Sales");
        if (range != "A1:B3") {
            throw new InvalidOperationException($"The Excel table used the unexpected range '{range}'.");
        }
        document.Save();
    }

    using ExcelDocument reopened = ExcelDocument.Load(path);
    if (reopened.Sheets.Count != 1 || reopened.Sheets[0].Name != "NativeAOT data") {
        throw new InvalidOperationException("The Excel round trip lost its worksheet.");
    }
    if (!reopened.Sheets[0].TryGetCellText(2, 1, out string region) || region != "North") {
        throw new InvalidOperationException("The Excel round trip lost its typed table data.");
    }
    AotSalesRow mappedRow = reopened.Sheets[0].RowsAs<AotSalesRow>(map => map
        .FromColumn<string>("Region", static (row, value) => { row.Region = value; return row; })
        .FromColumn<decimal>("Revenue", static (row, value) => { row.Revenue = value; return row; }))
        .First();
    if (mappedRow.Region != "North" || mappedRow.Revenue != 1250000M) {
        throw new InvalidOperationException("The AOT-safe typed-row mapping returned unexpected data.");
    }

    Console.WriteLine("PASS | Excel typed table create, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}

static async IAsyncEnumerable<SalesRow> CreateRowsAsync() {
    await Task.CompletedTask;
    yield return new SalesRow("North", 1250000M);
    yield return new SalesRow("South", 980000M);
}

internal readonly record struct SalesRow(string Region, decimal Revenue);

internal sealed class AotSalesRow {
    public string Region { get; set; } = string.Empty;
    public decimal Revenue { get; set; }
}
