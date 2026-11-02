using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNumericFormatNativeExportTests {
    [Fact]
    public void Native_reference_artifacts_match_their_hashes() {
        using var manifest = Manifest();
        foreach (JsonProperty artifact in manifest.RootElement.GetProperty("artifacts").EnumerateObject()) {
            byte[] bytes = File.ReadAllBytes(Corpus(artifact.Name));
            Assert.Equal(artifact.Value.GetProperty("bytes").GetInt64(), bytes.Length);
            Assert.Equal(artifact.Value.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Created_bytes_match_the_independently_decoded_Apple_accepted_profile(IWorkDocumentKind kind) {
        using var manifest = Manifest(); var family = Family(manifest.RootElement, kind);
        Assert.Equal(File.ReadAllBytes(Corpus(family.GetProperty("created").GetString()!)), CreateExample(kind, manifest.RootElement));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Created_and_Apple_edited_values_and_formats_survive_source_read_and_saved_destinations(IWorkDocumentKind kind, bool edited) {
        using var manifest = Manifest(); var root = manifest.RootElement; var family = Family(root, kind);
        var source = edited
            ? IWorkSourceDocument.Open(Corpus(family.GetProperty("edited").GetString()!))
            : IWorkSourceDocument.Open(CreateExample(kind, root));
        IWorkTable table = ReadTable(source, kind);
        Assert.Equal(20, table.RowCount); Assert.Equal(2, table.ColumnCount);
        foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray()) {
            int index = row.GetProperty("row").GetInt32();
            Assert.Equal(row.GetProperty("label").GetString(), table.GetCell(index, 1)!.Value);
            IWorkTableCell cell = table.GetCell(index, 2)!;
            Assert.Equal(IWorkCellKind.Number, cell.Kind);
            Assert.Equal(row.GetProperty(edited ? "editedValue" : "value").GetDouble(), cell.Value);
            IWorkNumberFormat format = Assert.IsType<IWorkNumberFormat>(cell.NumberFormat);
            JsonElement expected = row.GetProperty("format");
            Assert.Equal(expected.GetProperty("kind").GetString(), format.Kind.ToString());
            Assert.Equal(expected.GetProperty("decimalPlaces").GetInt32(), format.DecimalPlaces);
            Assert.Equal(expected.GetProperty("negativeStyle").GetInt32(), (int)format.NegativeStyle);
            Assert.Equal(expected.GetProperty("thousandsSeparator").GetBoolean(), format.ThousandsSeparator);
        }
        using var output = new MemoryStream();
        if (kind == IWorkDocumentKind.Numbers) {
            using var converted = source.ToExcelDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true });
            Assert.Contains(converted.Report.Diagnostics, item => item.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED");
            converted.Value.Save(output); output.Position = 0;
            using var saved = SpreadsheetDocument.Open(output, false); Assert.Empty(new OpenXmlValidator().Validate(saved));
            var cells = saved.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<S.Cell>().ToArray();
            var styles = saved.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            var codes = styles.NumberingFormats!.Elements<S.NumberingFormat>().ToDictionary(item => item.NumberFormatId!.Value,
                item => item.FormatCode!.Value);
            foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray()) {
                var cell = cells.Single(item => item.CellReference!.Value == "B" + row.GetProperty("row").GetInt32());
                Assert.True(cell.DataType == null || cell.DataType.Value == S.CellValues.Number);
                Assert.Equal(row.GetProperty(edited ? "editedValue" : "value").GetDouble(),
                    double.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture));
                uint formatId = styles.CellFormats!.Elements<S.CellFormat>().ElementAt((int)cell.StyleIndex!.Value).NumberFormatId!.Value;
                Assert.Equal(row.GetProperty("excelFormatCode").GetString(), codes[formatId]);
            }
            output.Position = 0; using var reopened = ExcelDocument.Load(output);
            Assert.Equal(edited ? -9876d : 1234.5, reopened.Sheets[0].CellAt(4, 2).GetValue<double>());
        } else if (kind == IWorkDocumentKind.Pages) {
            using var converted = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
            converted.RequireCompleteEditableReconstruction(); converted.Value.Save(output); output.Position = 0;
            using var saved = WordprocessingDocument.Open(output, false); Assert.Empty(new OpenXmlValidator().Validate(saved));
            var rows = saved.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single().Elements<W.TableRow>().ToArray();
            VerifyText(root, edited, rows.Select(row => row.Elements<W.TableCell>().Last().InnerText).ToArray());
            Assert.Contains(rows[6].Descendants<W.Color>(), color => color.Val?.Value == "FF0000");
        } else {
            using var converted = source.ToPowerPointDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
            converted.RequireCompleteEditableReconstruction(); converted.Value.Save(output); output.Position = 0;
            using var saved = PresentationDocument.Open(output, false); Assert.Empty(new OpenXmlValidator().Validate(saved));
            var rows = saved.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Table>().Single().Elements<A.TableRow>().ToArray();
            VerifyText(root, edited, rows.Select(row => row.Elements<A.TableCell>().Last().InnerText).ToArray());
            Assert.Contains(rows[6].Elements<A.TableCell>().Last().Descendants<A.RgbColorModelHex>(), color => color.Val?.Value == "FF0000");
        }
    }

    [Fact]
    public void Native_editable_exports_and_PDFs_retain_all_edited_numeric_values_and_displays() {
        using var manifest = Manifest(); var root = manifest.RootElement;
        foreach (JsonElement family in root.GetProperty("families").EnumerateArray()) {
            var pdf = PdfDocument.Load(new System.IO.MemoryStream(File.ReadAllBytes(Corpus(family.GetProperty("nativePdf").GetString()!))));
            Assert.Single(pdf.Pages); string text = Compact(pdf.ExtractText());
            foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray())
                Assert.Contains(Compact(row.GetProperty("label").GetString()! + row.GetProperty("nativeEditedDisplay").GetString()), text);
            string file = Corpus(family.GetProperty("nativeEditable").GetString()!);
            if (family.GetProperty("kind").GetString() == "Numbers") {
                using var workbook = ExcelDocument.Load(file);
                foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray())
                    Assert.Equal(row.GetProperty("editedValue").GetDouble(),
                        workbook.Sheets[0].CellAt(row.GetProperty("row").GetInt32() + 1, 3).GetValue<double>());
            } else {
                string[] values;
                if (file.EndsWith("docx", StringComparison.Ordinal)) {
                    using var doc = WordprocessingDocument.Open(file, false);
                    values = doc.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single().Elements<W.TableRow>()
                        .Select(row => row.Elements<W.TableCell>().Last().InnerText).ToArray();
                } else {
                    using var deck = PresentationDocument.Open(file, false);
                    values = deck.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Table>().Single().Elements<A.TableRow>()
                        .Select(row => row.Elements<A.TableCell>().Last().InnerText).ToArray();
                }
                Assert.Equal(20, values.Length);
                foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray())
                    Assert.Equal(Compact(row.GetProperty("nativeEditedDisplay").GetString()!), Compact(values[row.GetProperty("row").GetInt32() - 1]));
            }
        }
    }

    private static void VerifyText(JsonElement root, bool edited, string[] values) {
        Assert.Equal(20, values.Length);
        foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray())
            Assert.Equal(row.GetProperty(edited ? "invariantEditedDisplay" : "invariantDisplay").GetString(),
                values[row.GetProperty("row").GetInt32() - 1]);
    }

    private static IWorkTable ReadTable(IWorkSourceDocument source, IWorkDocumentKind kind) {
        if (kind == IWorkDocumentKind.Numbers) {
            var projection = source.ReadNumbers(); Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues);
            Assert.Empty(projection.SourceReferenceIssues); return Assert.Single(projection.Sheets[0].Tables);
        }
        if (kind == IWorkDocumentKind.Pages) {
            var projection = source.ReadPages(); Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues);
            Assert.Empty(projection.SourceReferenceIssues); return Assert.Single(projection.Tables);
        }
        var deck = source.ReadKeynote(); Assert.Empty(deck.Diagnostics); Assert.Empty(deck.SourceDeclarationIssues);
        Assert.Empty(deck.SourceReferenceIssues); return Assert.Single(deck.Slides[0].Tables);
    }

    private static byte[] CreateExample(IWorkDocumentKind kind, JsonElement root) {
        var data = IWorkTableData.Create(20, 2).SetText(1, 1, "Explicit numeric formats").SetText(1, 2, "Numeric values");
        foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray()) {
            var format = row.GetProperty("format"); int decimals = format.GetProperty("decimalPlaces").GetInt32();
            bool grouping = format.GetProperty("thousandsSeparator").GetBoolean();
            var negative = (IWorkNegativeNumberStyle)format.GetProperty("negativeStyle").GetInt32();
            IWorkNumberFormat number = format.GetProperty("kind").GetString() switch {
                "Number" => IWorkNumberFormat.CreateNumber(decimals, grouping, negative),
                "Percentage" => IWorkNumberFormat.CreatePercentage(decimals, grouping, negative),
                "Scientific" => IWorkNumberFormat.CreateScientific(decimals), _ => throw new InvalidDataException()
            };
            int index = row.GetProperty("row").GetInt32();
            data.SetText(index, 1, row.GetProperty("label").GetString()!).SetNumber(index, 2, row.GetProperty("value").GetDouble(), number);
        }
        if (kind == IWorkDocumentKind.Numbers) {
            var document = IWorkNumbersDocument.Create(); document.AddSheet("Formats").AddTable(data, "Formats", 20, 20, 280, 24);
            return document.SaveBytes();
        }
        if (kind == IWorkDocumentKind.Pages) {
            var document = IWorkPagesDocument.Create(); document.AddTable(data, "Formats", 234, 24, fontSizePoints: 10);
            return document.SaveBytes();
        }
        var deck = IWorkKeynoteDocument.Create(); deck.AddSlide().AddTable(data, 120, 40, 720, 480, "Formats", fontSizePoints: 12);
        return deck.SaveBytes();
    }

    private static JsonElement Family(JsonElement root, IWorkDocumentKind kind) => root.GetProperty("families").EnumerateArray()
        .Single(item => item.GetProperty("kind").GetString() == kind.ToString());
    private static string Compact(string text) => Regex.Replace(text, @"\s+", string.Empty);
    private static string Corpus(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
    private static JsonDocument Manifest() => JsonDocument.Parse(File.ReadAllText(Corpus("iwork-created-numeric-formats.json")));
}
