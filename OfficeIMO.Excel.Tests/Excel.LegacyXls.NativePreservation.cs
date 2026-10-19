using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyNativePreservationTests {
    private static string Fixture => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xls");

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public async Task Literal_text_edit_preserves_original_chart_and_compound_streams(bool asynchronous, bool continuedUnicode) {
        byte[] original = File.ReadAllBytes(Fixture);
        using var document = ExcelDocument.Load(new MemoryStream(original));
        string value = continuedUnicode ? string.Concat(Enumerable.Repeat("Łódź 🧪", 3000)) : "Edited inert preservation value";
        document.Sheets[0].CellValue(1, 1, value);
        using var output = new MemoryStream();
        if (asynchronous) await document.SaveAsync(output, ExcelFileFormat.Xls);
        else document.Save(output, ExcelFileFormat.Xls);
        byte[] saved = output.ToArray();
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? before, out string? error), error);
        Assert.True(OfficeCompoundFileReader.TryRead(saved, out OfficeCompoundFile? after, out error), error);
        Assert.Equal(before!.Streams.Keys.OrderBy(item => item), after!.Streams.Keys.OrderBy(item => item));
        foreach (var entry in before.Streams.Where(item => !item.Key.Equals("Workbook", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(entry.Value, after.Streams[entry.Key]);
        Assert.Contains(before.Streams.Keys, name => name.StartsWith("_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase));
        using var beforeModel = new MemoryStream(original); using var afterModel = new MemoryStream(saved);
        var beforeBook = LegacyXlsWorkbook.Load(beforeModel); var afterBook = LegacyXlsWorkbook.Load(afterModel);
        Assert.Equal(4, afterBook.ChartSheets.Count);
        AssertExtSstPointers(after.Streams["Workbook"]);
        for (int index = 0; index < beforeBook.ChartSheets.Count; index++)
            Assert.Equal(Substream(before.Streams["Workbook"], beforeBook.ChartSheets[index].StreamOffset), Substream(after.Streams["Workbook"], afterBook.ChartSheets[index].StreamOffset));
        using var reloaded = ExcelDocument.Load(new MemoryStream(saved));
        Assert.Equal(value, reloaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        document.Sheets[0].CellValue(1, 1, "Second preserving edit");
        using var second = ExcelDocument.Load(new MemoryStream(document.ToBytes(ExcelFileFormat.Xls)));
        Assert.Equal("Second preserving edit", second.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    [Fact]
    public void Existing_numeric_record_can_change_without_rebuilding_reference_tables() {
        var source = LegacyXlsWorkbook.Load(Fixture);
        var nativeSheet = source.Worksheets.First(sheet => sheet.Cells.Any(cell => cell.Kind == LegacyXlsCellValueKind.Number && !cell.IsFormula));
        var nativeCell = nativeSheet.Cells.First(cell => cell.Kind == LegacyXlsCellValueKind.Number && !cell.IsFormula);
        using var document = ExcelDocument.Load(Fixture);
        document.Sheets.First(sheet => sheet.Name == nativeSheet.Name).CellValue(nativeCell.Row, nativeCell.Column, 1.25D);
        using var saved = ExcelDocument.Load(new MemoryStream(document.ToBytes(ExcelFileFormat.Xls)));
        Assert.True(saved.Sheets.First(sheet => sheet.Name == nativeSheet.Name).TryGetCellText(nativeCell.Row, nativeCell.Column, out string text));
        Assert.Equal("1.25", text);
    }

    [Fact]
    public void Assigning_an_existing_plain_string_reuses_its_native_index() {
        using var document = ExcelDocument.Load(Fixture);
        var source = LegacyXlsWorkbook.Load(Fixture);
        string value = (string)source.Worksheets.SelectMany(sheet => sheet.Cells).First(cell => cell.Kind == LegacyXlsCellValueKind.Text && !cell.IsFormula && cell.TextFormattingRuns.Count == 0 && cell.Value is string text && text.Length > 0 && text != "Inert preservation control").Value!;
        document.Sheets[0].CellValue(1, 1, value);
        byte[] bytes = document.ToBytes(ExcelFileFormat.Xls);
        using var saved = ExcelDocument.Load(new MemoryStream(bytes));
        Assert.Equal(value, saved.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(Fixture), out OfficeCompoundFile? before, out string? error), error);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? after, out error), error);
        var oldSst = Assert.Single(BiffRecordReader.ReadRecords(before!.Streams["Workbook"], new()), record => record.Type == 0x00fc);
        var newSst = Assert.Single(BiffRecordReader.ReadRecords(after!.Streams["Workbook"], new()), record => record.Type == 0x00fc);
        Assert.Equal(BiffRecordReader.ReadUInt32(oldSst.Payload, 4), BiffRecordReader.ReadUInt32(newSst.Payload, 4));
    }

    [Fact]
    public void Encrypted_literal_edit_retains_VBA_streams_and_chart_sheet_records() {
        byte[] original = File.ReadAllBytes(Fixture);
        using var document = ExcelDocument.Load(Fixture);
        document.Sheets[0].CellValue(1, 1, "Encrypted preserving edit");
        using var output = new MemoryStream(); document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls);
        using var saved = ExcelDocument.LoadEncrypted(output, "test-password");
        Assert.Equal("Encrypted preserving edit", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(4, saved.LegacyXlsChartSheets.Count);
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? before, out string? error), error);
        Assert.True(OfficeCompoundFileReader.TryRead(output.ToArray(), out OfficeCompoundFile? after, out error), error);
        foreach (var entry in before!.Streams.Where(item => !item.Key.Equals("Workbook", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(entry.Value, after!.Streams[entry.Key]);
        output.Position = 0; Assert.ThrowsAny<Exception>(() => ExcelDocument.LoadEncrypted(output, "wrong-password"));
    }

    [Fact]
    public async Task Limits_and_cancellation_leave_the_destination_bytes_untouched() {
        using var document = ExcelDocument.Load(Fixture); document.Sheets[0].CellValue(1, 1, "Bounded edit");
        using var output = new MemoryStream(); output.WriteByte(42);
        Assert.Throws<InvalidDataException>(() => document.Save(output, ExcelFileFormat.Xls, new ExcelSaveOptions { MaxInMemoryPackageBytes = 1 }));
        Assert.Equal(new byte[] { 42 }, output.ToArray());
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.SaveAsync(output, ExcelFileFormat.Xls, cancellationToken: cancellation.Token));
        Assert.Equal(new byte[] { 42 }, output.ToArray());
    }

    [Fact]
    public void Shared_string_format_edits_are_encoded_without_reindexing_original_strings() {
        using var document = ExcelDocument.Load(new MemoryStream(File.ReadAllBytes(Fixture)));
        var cell = document.Sheets[0].WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().First(item => item.CellReference?.Value == "A1");
        int index = int.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture);
        var item = document.WorkbookPartRoot.SharedStringTablePart!.SharedStringTable!.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>().ElementAt(index);
        string text = item.InnerText; item.RemoveAllChildren();
        item.Append(new DocumentFormat.OpenXml.Spreadsheet.Run(new DocumentFormat.OpenXml.Spreadsheet.RunProperties(new DocumentFormat.OpenXml.Spreadsheet.Bold()), new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        using var output = new MemoryStream(); document.Save(output, ExcelFileFormat.Xls);
        using var reloaded = ExcelDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(text, reloaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.All(reloaded.Sheets[0].CellAt(1, 1).GetRichText(), run => Assert.True(run.Bold));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Native_same_path_file_save_keeps_the_binary_workbook_and_its_carriers(bool asynchronous) {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-Preserved-" + Guid.NewGuid().ToString("N") + ".xls");
        try {
            File.Copy(Fixture, path);
            using (var document = ExcelDocument.Load(path)) {
                document.Sheets[0].CellValue(1, 1, "First file edit");
                if (asynchronous) await document.SaveAsync(); else document.Save();
                document.Sheets[0].CellValue(1, 1, "Second file edit");
                if (asynchronous) await document.SaveAsync(); else document.Save();
            }
            using var saved = ExcelDocument.Load(path);
            Assert.Equal("Second file edit", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.Equal(4, saved.LegacyXlsChartSheets.Count);
            Assert.Contains(saved.LegacyXlsCompoundFeatures, feature => feature.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Theory]
    [InlineData(4, false)]
    [InlineData(5, false)]
    [InlineData(4, true)]
    [InlineData(5, true)]
    public void Pending_save_work_cannot_be_silently_skipped(int pendingWork, bool encrypted) {
        using var document = ExcelDocument.Load(Fixture);
        document.Sheets[0].CellValue(1, 1, "Pending work edit");
        switch (pendingWork) {
            case 4: document.SetCustomDocumentProperty("PreservationNote", "Requested metadata"); break;
            case 5: document.BuiltinDocumentProperties.Title = "Requested title"; break;
        }
        using var output = new MemoryStream(); output.WriteByte(42);
        Assert.Throws<NotSupportedException>(() => {
            if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls);
            else document.Save(output, ExcelFileFormat.Xls);
        });
        Assert.Equal(new byte[] { 42 }, output.ToArray());
    }

    [Fact]
    public void Short_native_SST_declines_preservation_without_breaking_salvage_load() {
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(Fixture), out OfficeCompoundFile? compound, out string? error), error);
        using var workbook = new MemoryStream();
        foreach (var record in BiffRecordReader.ReadRecords(compound!.Streams["Workbook"], new())) {
            byte[] payload = record.Type == 0x00fc ? record.Payload.Take(4).ToArray() : record.Payload;
            workbook.WriteByte((byte)record.Type); workbook.WriteByte((byte)(record.Type >> 8));
            workbook.WriteByte((byte)payload.Length); workbook.WriteByte((byte)(payload.Length >> 8));
            workbook.Write(payload, 0, payload.Length);
        }
        byte[] malformed = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Workbook"] = workbook.ToArray() });
        using var result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(malformed));
        Assert.True(result.HasDocument);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-SST-SHORT");
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    [InlineData(3, true)]
    public void Fractional_timestamp_edits_are_not_silently_discarded(int property, bool encrypted) {
        DateTime baseline = new(2026, 10, 19, 12, 30, 0, DateTimeKind.Utc);
        using var seed = ExcelDocument.Load(new MemoryStream(File.ReadAllBytes(Fixture)));
        seed.BuiltinDocumentProperties.Created = baseline;
        seed.BuiltinDocumentProperties.Modified = baseline;
        seed.BuiltinDocumentProperties.LastPrinted = baseline;
        seed.SetCustomDocumentProperty("ReviewedAt", baseline);
        var metadata = OfficeIMO.Excel.LegacyXls.Write.LegacyOlePropertySetWriter.CreateDocumentPropertyStreams(seed)
            .ToDictionary(item => item.Name, item => item.Bytes);
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(Fixture), out OfficeCompoundFile? compound, out string? error), error);
        using var document = ExcelDocument.Load(new MemoryStream(OfficeCompoundFileWriter.Rewrite(compound!, metadata)));
        Assert.Equal(baseline, document.CustomDocumentProperties["ReviewedAt"].Date);
        using var unchanged = new MemoryStream(); document.Save(unchanged, ExcelFileFormat.Xls);
        DateTime changed = baseline.AddTicks(1);
        switch (property) {
            case 0: document.BuiltinDocumentProperties.Created = changed; break;
            case 1: document.BuiltinDocumentProperties.Modified = changed; break;
            case 2: document.BuiltinDocumentProperties.LastPrinted = changed; break;
            case 3: document.CustomDocumentProperties["ReviewedAt"].Value = changed; break;
        }
        using var output = new MemoryStream(); output.WriteByte(42);
        Assert.Throws<NotSupportedException>(() => {
            if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls);
            else document.Save(output, ExcelFileFormat.Xls);
        });
        Assert.Equal(new byte[] { 42 }, output.ToArray());
    }

    private static void AssertExtSstPointers(byte[] bytes) {
        var records = BiffRecordReader.ReadRecords(bytes, new List<OfficeIMO.Excel.LegacyXls.Diagnostics.LegacyXlsImportDiagnostic>());
        var offsets = records.ToDictionary(item => item.Offset);
        var ext = Assert.Single(records, item => item.Type == 0x00ff);
        for (int index = 2; index < ext.Payload.Length; index += 8) {
            uint stringOffset = BiffRecordReader.ReadUInt32(ext.Payload, index);
            ushort relative = BiffRecordReader.ReadUInt16(ext.Payload, index + 4);
            Assert.True(offsets.TryGetValue(checked((int)(stringOffset - relative)), out var container));
            Assert.True(container.Type == 0x00fc || container.Type == 0x003c);
            Assert.InRange(relative, (ushort)4, checked((ushort)(container.Payload.Length + 2)));
        }
    }

    private static byte[] Substream(byte[] bytes, int offset) {
        int start = offset, depth = 0;
        do {
            ushort type = BiffRecordReader.ReadUInt16(bytes, offset), length = BiffRecordReader.ReadUInt16(bytes, offset + 2);
            if (type == 0x0809) depth++;
            if (type == 0x000a) depth--;
            offset += length + 4;
        } while (depth != 0);
        return bytes.Skip(start).Take(offset - start).ToArray();
    }
}
