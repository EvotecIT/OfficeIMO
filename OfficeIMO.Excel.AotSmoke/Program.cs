using OfficeIMO.Excel;

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".xlsx");
try {
    using (ExcelDocument document = ExcelDocument.Create(path)) {
        document.AddWorksheet("NativeAOT data");
        document.Save();
    }

    using ExcelDocument reopened = ExcelDocument.Load(path);
    if (reopened.Sheets.Count != 1 || reopened.Sheets[0].Name != "NativeAOT data") {
        throw new InvalidOperationException("The Excel round trip lost its worksheet.");
    }

    Console.WriteLine("PASS | Excel create, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}
