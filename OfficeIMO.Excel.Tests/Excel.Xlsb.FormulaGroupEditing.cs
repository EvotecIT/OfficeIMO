using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Write;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelXlsbFormulaGroupEditingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_arrays_keep_their_range_and_all_member_caches_during_repeated_saves(bool encrypted) {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddWorksheet("Data");
        sheet.CellValue(1, 1, 2D); sheet.CellValue(2, 1, 3D); sheet.SetArrayFormula("B1:B2", "SUM(A1:A2)");
        sheet.CellValue(1, 2, 5D); sheet.CellValue(2, 2, 5D);
        using var output = new MemoryStream();
        if (encrypted) seed.SaveEncrypted(output, "test-password", ExcelFileFormat.Xlsb); else seed.Save(output, ExcelFileFormat.Xlsb);
        using var document = encrypted ? ExcelDocument.LoadEncrypted(output, "test-password") : ExcelDocument.Load(output);
        var anchor = document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B1");
        Assert.Equal(CellFormulaValues.Array, anchor.CellFormula?.FormulaType?.Value); Assert.Equal("B1:B2", anchor.CellFormula?.Reference?.Value);
        document.Sheets[0].CellBold(2, 2);
        byte[] cleared = document.ToBytes(ExcelFileFormat.Xlsb, new() { ClearCachedFormulaResultsBeforeSave = true });
        var native = XlsbWorkbookReader.Load(cleared, new());
        Assert.Single(native.Sheets[0].FormulaGroups); Assert.All(native.Sheets[0].Cells.Where(cell => cell.Column == 2), cell => Assert.Equal(string.Empty, cell.Value));
        using var reloaded = ExcelDocument.Load(new MemoryStream(cleared));
        Assert.True(reloaded.Sheets[0].CellAt(2, 2).GetStyle().Bold);
        reloaded.Sheets[0].ClearArrayFormula("B2"); reloaded.Sheets[0].CellValue(1, 2, 99D);
        byte[] removed = reloaded.ToBytes(ExcelFileFormat.Xlsb);
        Assert.Empty(XlsbWorkbookReader.Load(removed, new()).Sheets[0].FormulaGroups);
        using var final = ExcelDocument.Load(new MemoryStream(removed)); Assert.Equal(99D, final.Sheets[0].CellAt(1, 2).GetValue<double>());
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void Disjoint_array_groups_keep_followers_and_unrelated_cells_when_caches_are_cleared(ExcelFileFormat format) {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddWorksheet("Arrays");
        var ranges = new List<(int Row, int Column)>();
        for (int group = 0; group < 12; group++) {
            int row = 1 + group / 4 * 100, column = 2 + group % 4 * 4;
            ranges.Add((row, column));
            sheet.SetArrayFormula(A1.CellReference(row, column) + ":" + A1.CellReference(row + 1, column), "2+3");
            sheet.CellValue(row, column, 5D); sheet.CellValue(row + 1, column, 5D);
        }
        sheet.CellValue(50, 20, 77D);
        using var document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(format)));
        document.Sheets[0].CellBold(202, 14);
        byte[] output = document.ToBytes(format, new() { ClearCachedFormulaResultsBeforeSave = true });
        using var copy = ExcelDocument.Load(new MemoryStream(output));
        var cells = copy.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        foreach (var position in ranges) {
            var owner = cells[A1.CellReference(position.Row, position.Column)];
            Assert.Equal(CellFormulaValues.Array, owner.CellFormula?.FormulaType?.Value);
            Assert.True(string.IsNullOrEmpty(owner.CellValue?.Text));
            Assert.True(string.IsNullOrEmpty(cells[A1.CellReference(position.Row + 1, position.Column)].CellValue?.Text));
        }
        Assert.True(copy.Sheets[0].CellAt(202, 14).GetStyle().Bold);
        Assert.Equal(77D, copy.Sheets[0].CellAt(50, 20).GetValue<double>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void First_off_origin_edit_updates_the_native_dimension_for_subsequent_saves(bool encryptedFirst) {
        using var seed = ExcelDocument.Create(); seed.AddWorksheet("Empty");
        using var document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(ExcelFileFormat.Xlsb)));
        document.Sheets[0].CellValue(7, 4, 12.5D);
        using var first = new MemoryStream();
        if (encryptedFirst) document.SaveEncrypted(first, "test-password", ExcelFileFormat.Xlsb); else document.Save(first, ExcelFileFormat.Xlsb);
        document.Sheets[0].CellValue(8, 4, 13.5D);
        using var second = new MemoryStream();
        if (encryptedFirst) document.Save(second, ExcelFileFormat.Xlsb); else document.SaveEncrypted(second, "test-password", ExcelFileFormat.Xlsb);
        using var copy = encryptedFirst ? ExcelDocument.Load(second) : ExcelDocument.LoadEncrypted(second, "test-password");
        Assert.Equal(12.5D, copy.Sheets[0].CellAt(7, 4).GetValue<double>());
        Assert.Equal(13.5D, copy.Sheets[0].CellAt(8, 4).GetValue<double>());
    }

    [Fact]
    public void Replacing_the_owner_of_a_native_shared_group_preserves_other_relative_formulas() {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddWorksheet("Data");
        for (int row = 1; row <= 3; row++) { sheet.CellValue(row, 1, row * 10D); sheet.CellFormula(row, 2, "A" + row + "+5"); }
        sheet.RecalculateSupportedFormulas(); byte[] bytes = seed.ToBytes(ExcelFileFormat.Xlsb);
        using var input = new MemoryStream(bytes); using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        using var changed = new MemoryStream();
        using (var target = new ZipArchive(changed, ZipArchiveMode.Create, true)) foreach (var entry in archive.Entries) {
            var output = target.CreateEntry(entry.FullName); using var destination = output.Open(); using var source = entry.Open();
            if (!entry.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal) || !entry.FullName.EndsWith(".bin", StringComparison.Ordinal)) { source.CopyTo(destination); continue; }
            var records = XlsbRecordReader.ReadAll(source); bool declaration = false; uint currentRow = 0;
            foreach (var record in records) {
                if (record.Type == 0) currentRow = XlsbFormulaGroup.Read(record.Data, 0);
                if (record.Type != 9) { XlsbRecordWriter.Write(destination, record.Type, record.Data); continue; }
                byte[] payload = new byte[35]; Buffer.BlockCopy(record.Data, 0, payload, 0, 16);
                payload[18] = 5; payload[22] = 1; // PtgExp
                XlsbFormulaGroup.Put(payload, 23, currentRow); XlsbFormulaGroup.Put(payload, 27, 4); XlsbFormulaGroup.Put(payload, 31, 1);
                XlsbRecordWriter.Write(destination, 9, payload);
                if (!declaration) {
                    byte[] definition = new byte[36]; XlsbFormulaGroup.Put(definition, 4, 2); XlsbFormulaGroup.Put(definition, 8, 1); XlsbFormulaGroup.Put(definition, 12, 1);
                    XlsbFormulaGroup.Put(definition, 16, 12); definition[20] = 0x4c; definition[25] = 0xff; definition[26] = 0xff; // RefNV A(relative -1), row(relative 0)
                    definition[27] = 0x1e; definition[28] = 5; definition[30] = 0x03;
                    // 7-byte RefN + 3-byte integer + Add = 11 token bytes.
                    Array.Resize(ref definition, 35); XlsbFormulaGroup.Put(definition, 16, 11);
                    XlsbRecordWriter.Write(destination, 427, definition); declaration = true;
                }
            }
        }
        using var document = ExcelDocument.Load(new MemoryStream(changed.ToArray()));
        document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B1").CellFormula = null;
        document.Sheets[0].CellValue(1, 2, 100D); document.Sheets[0].CellBold(3, 2);
        byte[] saved = document.ToBytes(ExcelFileFormat.Xlsb); var native = XlsbWorkbookReader.Load(saved, new());
        Assert.Empty(native.Sheets[0].FormulaGroups);
        Assert.Equal("A2+5", native.Sheets[0].Cells.Single(cell => cell.Row == 2 && cell.Column == 2).FormulaText);
        Assert.Equal("A3+5", native.Sheets[0].Cells.Single(cell => cell.Row == 3 && cell.Column == 2).FormulaText);
        using var reloaded = ExcelDocument.Load(new MemoryStream(saved)); reloaded.Sheets[0].RecalculateSupportedFormulas();
        Assert.Equal(25D, reloaded.Sheets[0].CellAt(2, 2).GetValue<double>()); Assert.Equal(35D, reloaded.Sheets[0].CellAt(3, 2).GetValue<double>());
    }
    [Fact]
    public void Repeated_formula_edits_reuse_the_native_self_reference_table() {
        using var seed = ExcelDocument.Create(); seed.AddWorksheet("Data").CellValue(1, 1, 5D);
        byte[] bytes = seed.ToBytes(ExcelFileFormat.Xlsb); int links = 0, references = 0;
        for (int iteration = 0; iteration < 5; iteration++) {
            using var document = ExcelDocument.Load(new MemoryStream(bytes));
            document.Sheets[0].CellFormula(2 + iteration, 1, "'Data'!A1+1"); document.Sheets[0].RecalculateSupportedFormulas();
            bytes = document.ToBytes(ExcelFileFormat.Xlsb);
            var native = XlsbWorkbookReader.Load(bytes);
            if (iteration == 0) { links = native.SelfSupportingLinks.Count; references = native.ExternalSheetReferences.Count; }
            else { Assert.Equal(links, native.SelfSupportingLinks.Count); Assert.Equal(references, native.ExternalSheetReferences.Count); }
            using var loaded = ExcelDocument.Load(new MemoryStream(bytes));
            Assert.Equal(6D, loaded.Sheets[0].CellAt(2 + iteration, 1).GetValue<double>());
        }
    }

    [Fact]
    public void Native_sheet_selection_respects_the_explicit_active_sheet_and_does_not_group_new_tabs() {
        using var seed = ExcelDocument.Create(); seed.AddWorksheet("First"); seed.AddWorksheet("Second"); seed.SetActiveWorksheet("Second");
        byte[] original = seed.ToBytes(ExcelFileFormat.Xlsb);
        Assert.False(IsSelected(original, "First")); Assert.True(IsSelected(original, "Second"));
        using var document = ExcelDocument.Load(new MemoryStream(original)); document.AddWorksheet("Added").CellValue(1, 1, "New tab");
        byte[] saved = document.ToBytes(ExcelFileFormat.Xlsb);
        Assert.False(IsSelected(saved, "First")); Assert.True(IsSelected(saved, "Second")); Assert.False(IsSelected(saved, "Added"));
        static bool IsSelected(byte[] bytes, string name) {
            var native = XlsbWorkbookReader.Load(bytes); var sheet = native.Sheets.Single(sheet => sheet.Name == name);
            using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            using var content = archive.GetEntry(sheet.PartName!.TrimStart('/'))!.Open();
            var view = Assert.Single(XlsbRecordReader.ReadAll(content), record => record.Type == 137);
            return (view.Data[0] & 0x40) != 0;
        }
    }

}
