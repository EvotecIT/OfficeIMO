using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkEverydayNativeExportTests {
    [Fact]
    public void Everyday_Pages_body_tables_and_section_text_match_independent_Apple_export() {
        VerifyManifest("pages-everyday-v15.4.json");
        using var result = WordIWorkConverter.ConvertPagesToWordResult(Corpus("picodocs/sample-v14.4.pages"),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.True(result.Report.IsPartialEditableReconstruction);
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var actual = WordprocessingDocument.Open(saved, false);
        using var apple = WordprocessingDocument.Open(Corpus("native-exports/pages-everyday-v15.4.docx"), false);
        W.Body expected = apple.MainDocumentPart!.Document!.Body!;
        W.Body body = actual.MainDocumentPart!.Document!.Body!;
        // Empty structural section paragraphs may differ; meaningful body order must not.
        Assert.Equal(expected.Elements<W.Paragraph>().Select(Text).Where(t => t.Length > 0),
            body.Elements<W.Paragraph>().Select(Text).Where(t => t.Length > 0));
        W.Table[] expectedTables = expected.Elements<W.Table>().ToArray();
        W.Table[] tables = body.Elements<W.Table>().ToArray();
        Assert.Equal(3, tables.Length);
        for (int index = 0; index < 2; index++)
            Assert.Equal(expectedTables[index].Descendants<W.TableCell>().Select(Text),
                tables[index].Descendants<W.TableCell>().Select(Text));
        // These seven cells contain date/currency display strings with unqualified locale rendering.
        int[] localeCells = [6, 8, 9, 11, 13, 14, 19];
        string[] nativeCells = expectedTables[2].Descendants<W.TableCell>().Select(Text).ToArray();
        string[] outputCells = tables[2].Descendants<W.TableCell>().Select(Text).ToArray();
        Assert.Equal(nativeCells.Length, outputCells.Length);
        for (int index = 0; index < nativeCells.Length; index++)
            if (!localeCells.Contains(index)) Assert.Equal(nativeCells[index], outputCells[index]);
        Assert.Equal(apple.MainDocumentPart.HeaderParts.Select(p => Text(p.Header!)).OrderBy(t => t),
            actual.MainDocumentPart.HeaderParts.Select(p => Text(p.Header!)).OrderBy(t => t));
        Assert.Equal(apple.MainDocumentPart.FooterParts.Select(p => Text(p.Footer!)).OrderBy(t => t),
            actual.MainDocumentPart.FooterParts.Select(p => Text(p.Footer!)).OrderBy(t => t));
        Assert.Equal(expected.Descendants<W.Drawing>().Count(), body.Descendants<W.Drawing>().Count());
    }

    [Fact]
    public void Cross_table_Numbers_values_and_formula_caches_match_all_Apple_worksheets() {
        VerifyManifest("numbers-cross-table-v14.5.json");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(Corpus("numbers-parser/cross-table-formulas.numbers"),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Contains(result.Report.FidelityDiagnostics, d => d.Code == "IWORK_NUMBERS_ERROR_VALUE_APPROXIMATED"
            && d.LossKind == OfficeConversionLossKind.Approximation);
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var actual = SpreadsheetDocument.Open(saved, false);
        using var apple = SpreadsheetDocument.Open(Corpus("native-exports/numbers-cross-table-v14.5.xlsx"), false);
        WorkbookPart reference = apple.WorkbookPart!;
        WorkbookPart output = actual.WorkbookPart!;
        var nativeSheets = reference.Workbook!.Sheets!.Elements<Sheet>().ToArray();
        var outputSheets = output.Workbook!.Sheets!.Elements<Sheet>().ToArray();
        Assert.Equal(14, nativeSheets.Length);
        Assert.Equal(nativeSheets.Select(s => s.Name!.Value).OrderBy(s => s), outputSheets.Select(s => s.Name!.Value).OrderBy(s => s));
        int verified = 0, formulaCaches = 0;
        foreach (Sheet sheet in nativeSheets) {
            Sheet destination = outputSheets.Single(s => s.Name!.Value == sheet.Name!.Value);
            var nativePart = (WorksheetPart)reference.GetPartById(sheet.Id!.Value!);
            var outputPart = (WorksheetPart)output.GetPartById(destination.Id!.Value!);
            var cells = outputPart.Worksheet!.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
            foreach (Cell nativeCell in nativePart.Worksheet!.Descendants<Cell>()) {
                string nativeAddress = nativeCell.CellReference!.Value!;
                uint row = uint.Parse(new string(nativeAddress.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
                if (row == 1) continue; // Apple's added table-title row.
                string address = new string(nativeAddress.Where(char.IsLetter).ToArray()) + (row - 1).ToString(CultureInfo.InvariantCulture);
                if (nativeCell.CellFormula?.Text == "SUM(#REF!)") continue; // Apple stores no cache; generic source error is reported above.
                (string kind, string? value) expected = Value(reference, nativeCell);
                if (expected.value is null) continue; // Formatting-only sparse cells.
                Assert.True(cells.TryGetValue(address, out Cell? cell), sheet.Name + "!" + address);
                var actualValue = Value(output, cell!);
                Assert.Equal(expected.kind, actualValue.kind);
                if (expected.kind == "number") Assert.Equal(double.Parse(expected.value, CultureInfo.InvariantCulture),
                    double.Parse(actualValue.value!, CultureInfo.InvariantCulture), 12);
                else Assert.Equal(expected.value, actualValue.value);
                Assert.Equal(nativeCell.CellFormula is not null, cell!.CellFormula is not null);
                if (nativeCell.CellFormula is not null) formulaCaches++;
                verified++;
            }
        }
        Assert.Equal(529, verified);
        Assert.Equal(109, formulaCaches);
    }

    [Fact]
    public void Uniform_formula_fonts_alignment_and_caches_match_independent_Numbers_export() {
        VerifyManifest("numbers-formula-formatting-v14.5.json");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(Corpus("native-exports/numbers-formula-formatting-v14.5.numbers"),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var actual = SpreadsheetDocument.Open(saved, false);
        using var apple = SpreadsheetDocument.Open(Corpus("native-exports/numbers-formula-formatting-v14.5.xlsx"), false);
        WorkbookPart reference = apple.WorkbookPart!, output = actual.WorkbookPart!;
        foreach ((string nativeAddress, string address, string expression) in new[] {
            ("A10", "A9", "\"\"\"\"&A4&\"\"\" \"\"\"&B4&\"\"\"\""),
            ("A11", "A10", "A5&\" \"&B5&\" melon\"") }) {
            Cell nativeCell = FormulaCell(reference, nativeAddress), cell = FormulaCell(output, address);
            Assert.Equal(Value(reference, nativeCell), Value(output, cell));
            Assert.NotNull(nativeCell.CellFormula);
            Assert.Equal(expression, cell.CellFormula!.Text);
            CellFormat expectedFormat = Format(reference, nativeCell), format = Format(output, cell);
            Assert.Equal(expectedFormat.Alignment!.Horizontal!.Value, format.Alignment!.Horizontal!.Value);
            Font expectedFont = Font(reference, expectedFormat), font = Font(output, format);
            Assert.Equal(Enabled(expectedFont.Bold), Enabled(font.Bold));
            Assert.Equal(Enabled(expectedFont.Italic), Enabled(font.Italic));
            Assert.Equal(expectedFont.Underline is not null, font.Underline is not null);
            Assert.Equal(Enabled(expectedFont.Strike), Enabled(font.Strike));
            Assert.Equal(expectedFont.FontSize!.Val!.Value, font.FontSize!.Val!.Value);
            Assert.Equal("Helvetica Neue", expectedFont.FontName!.Val!.Value);
            Assert.Equal("HelveticaNeue", font.FontName!.Val!.Value);
            Assert.Equal(8u, expectedFont.Color!.Indexed!.Value); // Native indexed black.
            Assert.Equal("FF000000", font.Color!.Rgb!.Value);
        }
    }

    private static Cell FormulaCell(WorkbookPart workbook, string address) {
        Sheet sheet = workbook.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name!.Value == "Main Sheet - Formula Tests");
        return ((WorksheetPart)workbook.GetPartById(sheet.Id!.Value!)).Worksheet!.Descendants<Cell>()
            .Single(c => c.CellReference!.Value == address);
    }

    private static CellFormat Format(WorkbookPart workbook, Cell cell) =>
        workbook.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
    private static Font Font(WorkbookPart workbook, CellFormat format) =>
        workbook.WorkbookStylesPart!.Stylesheet!.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value);
    private static bool Enabled(BooleanPropertyType? property) => property is not null && (property.Val?.Value ?? true);

    private static (string kind, string? value) Value(WorkbookPart workbook, Cell cell) {
        if (cell.DataType?.Value == CellValues.SharedString)
            return ("text", workbook.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>()
                .ElementAt(int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture)).InnerText);
        if (cell.DataType?.Value == CellValues.InlineString) return ("text", cell.InlineString?.InnerText);
        if (cell.DataType?.Value == CellValues.String) return ("text", cell.CellValue?.Text);
        if (cell.DataType?.Value == CellValues.Boolean) return ("boolean", cell.CellValue?.Text);
        if (cell.DataType?.Value == CellValues.Error) return ("error", cell.CellValue?.Text);
        return ("number", cell.CellValue?.Text);
    }

    private static string Text(DocumentFormat.OpenXml.OpenXmlElement element) => string.Concat(element.Descendants<W.Text>().Select(t => t.Text));

    private static void VerifyManifest(string name) {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Corpus("native-exports/" + name)));
        JsonElement root = manifest.RootElement;
        Assert.Equal(root.GetProperty("sourceSha256").GetString(), Hash(Corpus(root.GetProperty("source").GetString()!)));
        foreach (JsonElement artifact in root.GetProperty("artifacts").EnumerateArray())
            Assert.Equal(artifact.GetProperty("sha256").GetString(), Hash(Corpus(artifact.GetProperty("path").GetString()!)));
    }

    private static string Corpus(string path) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", path);
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
