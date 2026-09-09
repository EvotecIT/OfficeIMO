using System.Text;
using System.Xml;
using OfficeIMO.AI;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Reader;

internal static class ArtifactWriter {
    public static async Task SaveAsync(string output, OfficeAiDocument document, OfficeAiResult result, CancellationToken cancellationToken = default) {
        var artifacts = await PrepareAsync(document, result, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        Directory.CreateDirectory(output);
        foreach (var artifact in artifacts)
            await WriteNewAsync(Path.Combine(output, artifact.Name), artifact.Bytes, cancellationToken);
    }

    private static async Task<IReadOnlyList<(string Name, byte[] Bytes)>> PrepareAsync(
        OfficeAiDocument document, OfficeAiResult result, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var tables = new List<(string Name, IReadOnlyList<string> Columns, IReadOnlyList<IReadOnlyList<string>> Rows)>();
        if (result.Fields.Count > 0) tables.Add(("Fields", new[] { "Name", "Status", "Raw value", "Normalized value", "Evidence" },
            result.Fields.Select(field => {
                cancellationToken.ThrowIfCancellationRequested();
                return (IReadOnlyList<string>)new[] { field.Name, field.Status.ToString(), field.RawValue ?? "",
                    field.NormalizedValue ?? "", string.Join(", ", field.Citations.Select(citation => citation.EvidenceId)) };
            }).ToArray()));
        for (int index = 0; index < result.Tables.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            tables.Add(("Table " + (index + 1), result.Tables[index].Table.Columns, result.Tables[index].Table.Rows));
        }
        // JSON strings are bounded in Unicode scalars; the XLSX cell contract counts UTF-16 code units.
        // Check every exported value before writing any artifacts, including field/evidence summaries and headings.
        foreach (var table in tables)
            foreach (string value in table.Columns.Concat(table.Rows.SelectMany(row => row))) {
                cancellationToken.ThrowIfCancellationRequested();
                if (value.Length > 32_767)
                    throw new NotSupportedException("XLSX export requires every cell to fit Excel's 32,767 UTF-16 code-unit limit. No artifacts were written.");
                try { XmlConvert.VerifyXmlChars(value); }
                catch (XmlException) {
                    throw new NotSupportedException("XLSX export requires XML-compatible cell characters. No artifacts were written.");
                }
            }
        cancellationToken.ThrowIfCancellationRequested();
        var artifacts = new List<(string Name, byte[] Bytes)> {
            ("report.json", Encoding.UTF8.GetBytes(OfficeAiArtifacts.SerializeReport(document, result)))
        };
        if (result.Operation == OfficeAiOperation.Parse) {
            string readback = OfficeDocumentReadResultJson.Serialize(OfficeAiArtifacts.CreateProposedReadResult(document, result), indented: true);
            _ = OfficeDocumentReadResultJson.Deserialize(readback);
            artifacts.Add(("proposed-reader.json", Encoding.UTF8.GetBytes(readback)));
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (tables.Count == 0) return artifacts;
        using var workbookBytes = new MemoryStream();
        using (var workbook = ExcelDocument.Create(workbookBytes)) {
            foreach (var table in tables) {
                cancellationToken.ThrowIfCancellationRequested();
                var csv = new CsvDocument().WithHeader(table.Columns.ToArray());
                var sheet = workbook.AddWorksheet(table.Name);
                for (int column = 0; column < table.Columns.Count; column++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    sheet.CellValue(1, column + 1, table.Columns[column]);
                }
                for (int row = 0; row < table.Rows.Count; row++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    csv.AddRow(table.Rows[row].Cast<object>().ToArray());
                    for (int column = 0; column < table.Columns.Count; column++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        sheet.CellValue(row + 2, column + 1, table.Rows[row][column]);
                    }
                }
                using var csvBytes = new MemoryStream();
                await csv.SaveAsync(csvBytes, new CsvSaveOptions { IncludeHeader = true, FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape }, cancellationToken);
                artifacts.Add((table.Name.Replace(' ', '-') + ".csv", csvBytes.ToArray()));
            }
            await workbook.SaveAsync(cancellationToken);
        }
        byte[] workbookArtifact = workbookBytes.ToArray();
        using var readbackStream = new MemoryStream(workbookArtifact, writable: false);
        using var reopened = ExcelDocument.Load(readbackStream);
        cancellationToken.ThrowIfCancellationRequested();
        if (reopened.Sheets.Count != tables.Count) throw new InvalidDataException("Workbook readback did not preserve its worksheets.");
        for (int index = 0; index < tables.Count; index++) {
            var table = tables[index];
            var sheet = reopened.Sheets[index];
            for (int column = 0; column < table.Columns.Count; column++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (sheet.CellAt(1, column + 1).GetValue<string>() != table.Columns[column])
                    throw new InvalidDataException("Workbook readback changed a column heading.");
                for (int row = 0; row < table.Rows.Count; row++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if ((sheet.CellAt(row + 2, column + 1).GetValue<string>() ?? "") != table.Rows[row][column])
                        throw new InvalidDataException("Workbook readback changed an extracted value.");
                }
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        artifacts.Add(("extraction.xlsx", workbookArtifact));
        return artifacts;
    }

    private static async Task WriteNewAsync(string path, byte[] bytes, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        await using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, FileOptions.Asynchronous);
        await stream.WriteAsync(bytes, cancellationToken);
        await stream.FlushAsync(cancellationToken);
    }
}
