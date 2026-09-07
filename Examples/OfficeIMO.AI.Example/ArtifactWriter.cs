using System.Text;
using OfficeIMO.AI;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Reader;

internal static class ArtifactWriter {
    public static void Save(string output, OfficeAiDocument document, OfficeAiResult result) {
        Directory.CreateDirectory(output);
        WriteNew(Path.Combine(output, "report.json"), Encoding.UTF8.GetBytes(OfficeAiArtifacts.SerializeReport(document, result)));
        if (result.Operation == OfficeAiOperation.Parse) {
            string readback = OfficeDocumentReadResultJson.Serialize(OfficeAiArtifacts.CreateProposedReadResult(document, result), indented: true);
            WriteNew(Path.Combine(output, "proposed-reader.json"), Encoding.UTF8.GetBytes(readback));
            _ = OfficeDocumentReadResultJson.Deserialize(File.ReadAllText(Path.Combine(output, "proposed-reader.json")));
        }
        var tables = new List<(string Name, IReadOnlyList<string> Columns, IReadOnlyList<IReadOnlyList<string>> Rows)>();
        if (result.Fields.Count > 0) tables.Add(("Fields", new[] { "Name", "Status", "Raw value", "Normalized value", "Evidence" },
            result.Fields.Select(field => (IReadOnlyList<string>)new[] { field.Name, field.Status.ToString(), field.RawValue ?? "",
                field.NormalizedValue ?? "", string.Join(", ", field.Citations.Select(citation => citation.EvidenceId)) }).ToArray()));
        for (int index = 0; index < result.Tables.Count; index++) tables.Add(("Table " + (index + 1), result.Tables[index].Table.Columns, result.Tables[index].Table.Rows));
        if (tables.Count == 0) return;
        using var workbookBytes = new MemoryStream();
        using (var workbook = ExcelDocument.Create(workbookBytes)) {
            foreach (var table in tables) {
                var csv = new CsvDocument().WithHeader(table.Columns.ToArray());
                var sheet = workbook.AddWorksheet(table.Name);
                for (int column = 0; column < table.Columns.Count; column++) sheet.CellValue(1, column + 1, table.Columns[column]);
                for (int row = 0; row < table.Rows.Count; row++) {
                    csv.AddRow(table.Rows[row].Cast<object>().ToArray());
                    for (int column = 0; column < table.Columns.Count; column++) sheet.CellValue(row + 2, column + 1, table.Rows[row][column]);
                }
                using var csvBytes = new MemoryStream();
                csv.Save(csvBytes, new CsvSaveOptions { IncludeHeader = true, FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape });
                WriteNew(Path.Combine(output, table.Name.Replace(' ', '-') + ".csv"), csvBytes.ToArray());
            }
            workbook.Save();
        }
        string workbookPath = Path.Combine(output, "extraction.xlsx");
        WriteNew(workbookPath, workbookBytes.ToArray());
        using var reopened = ExcelDocument.Load(workbookPath);
        if (reopened.Sheets.Count != tables.Count) throw new InvalidDataException("Workbook readback did not preserve its worksheets.");
        for (int index = 0; index < tables.Count; index++) {
            var table = tables[index];
            var sheet = reopened.Sheets[index];
            for (int column = 0; column < table.Columns.Count; column++) {
                if (sheet.CellAt(1, column + 1).GetValue<string>() != table.Columns[column])
                    throw new InvalidDataException("Workbook readback changed a column heading.");
                for (int row = 0; row < table.Rows.Count; row++)
                    if ((sheet.CellAt(row + 2, column + 1).GetValue<string>() ?? "") != table.Rows[row][column])
                        throw new InvalidDataException("Workbook readback changed an extracted value.");
            }
        }
    }

    private static void WriteNew(string path, byte[] bytes) {
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        stream.Write(bytes);
    }
}
