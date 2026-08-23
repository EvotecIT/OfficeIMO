namespace OfficeIMO.OpenDocument.Benchmarks.Comparisons;

internal static class OdsComparisonWorkflows {
    internal static byte[] CreateOfficeIMO(OdsComparisonScale scale) {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        for (var row = 0; row < scale.Rows; row++) {
            for (var column = 0; column < scale.Columns; column++) {
                sheet.Cell(row, column).SetString(OdsComparisonCorpus.Cell(row, column));
            }
        }
        return document.ToBytes();
    }

    internal static async Task<byte[]> CreateOpenStandardLibrary(OdsComparisonScale scale) {
        await using var spreadsheet = new global::OoxSpreadsheet.Spreadsheet();
        OslSpreadsheet.Models.oSpreadsheet sheet = spreadsheet.Workbook.AddSheet("Data");
        for (var row = 0; row < scale.Rows; row++) {
            for (var column = 0; column < scale.Columns; column++) {
                sheet.AddCell(row + 1, column + 1, OdsComparisonCorpus.Cell(row, column));
            }
        }
        return await spreadsheet.GenerateOdsFileAsync().ConfigureAwait(false);
    }

    internal static long ReadOfficeIMO(byte[] package) {
        OdsDocument document = OdsDocument.Load(new MemoryStream(package, writable: false));
        OdsSheet sheet = document.Sheets.Single();
        long checksum = 0;
        foreach (OdsRowRun row in sheet.RowRuns) {
            foreach (OdsCellRun cell in row.CellRuns) {
                checksum += checked(row.RepeatCount * cell.RepeatCount * cell.Value.ToString().Length);
            }
        }
        return checksum;
    }

    internal static async Task<long> ReadOpenStandardLibrary(byte[] package) {
        await using var spreadsheet = new global::OoxSpreadsheet.Spreadsheet();
        OslSpreadsheet.Models.oWorkbook workbook = await spreadsheet.ImportOdsFileAsync(package).ConfigureAwait(false);
        return workbook.Sheets.Single().Cells.Sum(cell => (long)cell.Value.Length);
    }
}
