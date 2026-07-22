using System.Data;
using OfficeIMO.Excel;

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".xlsx");
try {
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

    Console.WriteLine("PASS | Excel typed table create, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}
