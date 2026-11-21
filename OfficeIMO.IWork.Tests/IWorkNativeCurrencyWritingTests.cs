using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNativeCurrencyWritingTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, "pages", "docx")]
    [InlineData(IWorkDocumentKind.Numbers, "numbers", "xlsx")]
    [InlineData(IWorkDocumentKind.Keynote, "key", "pptx")]
    public void Created_and_Apple_saved_tables_retain_currency_selections_and_values(
        IWorkDocumentKind kind, string nativeExtension, string exportExtension) {
        using var manifest = JsonDocument.Parse(File.ReadAllBytes(Corpus("manifest.json")));
        JsonElement hashes = manifest.RootElement.GetProperty("artifacts");
        string createdName = "currency-created." + nativeExtension;
        string editedName = "currency-edited." + nativeExtension;
        string exportBase = "edited-" + nativeExtension;
        byte[] created = Create(kind);
        Assert.Equal(File.ReadAllBytes(Corpus(createdName)), created);
        foreach (string name in new[] {
            createdName, editedName, "currency-created-" + nativeExtension + ".json",
            exportBase + ".json", exportBase + "." + exportExtension, exportBase + ".pdf"
        }) {
            Assert.Equal(hashes.GetProperty(name).GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Corpus(name)))).ToLowerInvariant());
        }

        VerifyReadback(kind, created, "currency-created-" + nativeExtension + ".json", -12.5);
        VerifyReadback(kind, File.ReadAllBytes(Corpus(editedName)), exportBase + ".json", -34.5);
        VerifyExport(kind, Corpus(exportBase + "." + exportExtension));
        var pdf = PdfDocument.Load(new MemoryStream(File.ReadAllBytes(Corpus(exportBase + ".pdf"))));
        Assert.Single(pdf.Pages);
        string visibleText = Compact(pdf.ExtractText());
        foreach (string amount in new[] { "-34,50zł", "12,50USD", "(1234JPY)", "-7,25USD" })
            Assert.Contains(amount, visibleText);
    }

    private static void VerifyReadback(IWorkDocumentKind kind, byte[] sourceBytes, string decodeName, double firstAmount) {
        using var independent = JsonDocument.Parse(File.ReadAllBytes(Corpus(decodeName)));
        JsonElement reference = Assert.Single(independent.RootElement.GetProperty("tables").EnumerateArray());
        var source = IWorkSourceDocument.Load(sourceBytes);
        IWorkTable table = kind switch {
            IWorkDocumentKind.Pages => Assert.Single(source.ReadPages().Tables),
            IWorkDocumentKind.Numbers => Assert.Single(Assert.Single(source.ReadNumbers().Sheets).Tables),
            _ => Assert.Single(Assert.Single(source.ReadKeynote().Slides).Tables)
        };
        Assert.Equal(5, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(table.RowCount, reference.GetProperty("rows").GetInt32());
        Assert.Equal(table.ColumnCount, reference.GetProperty("columns").GetInt32());
        JsonElement[] independentCells = reference.GetProperty("numericCells").EnumerateArray().ToArray();
        Assert.Equal(new[] { firstAmount, 12.5, -1234, -7.25 },
            independentCells.Select(cell => cell.GetProperty("value").GetDouble()));
        Assert.Equal(new[] { "PLN", "USD", "JPY", "USD" }, independentCells.Select(
            cell => cell.GetProperty("numberFormat").GetProperty("currencyCode").GetString()));
        Assert.Equal(new[] { 2, 3, 4, 5 }, independentCells.Select(cell => cell.GetProperty("row").GetInt32()));

        foreach (JsonElement expected in independentCells) {
            Assert.True(expected.GetProperty("activeNumericFormat").GetBoolean());
            Assert.Equal(2, expected.GetProperty("suggestedType").GetInt32());
            Assert.Equal(2, expected.GetProperty("column").GetInt32());
            IWorkTableCell cell = table.GetCell(expected.GetProperty("row").GetInt32(), 2)!;
            Assert.Equal(IWorkCellKind.Number, cell.Kind);
            Assert.Equal(expected.GetProperty("value").GetDouble(), cell.Value);
            JsonElement format = expected.GetProperty("numberFormat");
            var selected = Assert.IsType<IWorkNumberFormat>(cell.NumberFormat);
            Assert.Equal(IWorkNumberFormatKind.Currency, selected.Kind);
            Assert.Equal(format.GetProperty("currencyCode").GetString(), selected.CurrencyCode);
            Assert.Equal(format.GetProperty("decimalPlaces").GetInt32(), selected.DecimalPlaces);
            Assert.Equal(format.GetProperty("negativeStyle").GetInt32(), (int)selected.NegativeStyle);
            Assert.Equal(format.GetProperty("thousandsSeparator").GetBoolean(), selected.ThousandsSeparator);
            Assert.False(selected.UseAccountingStyle);
        }
    }

    private static void VerifyExport(IWorkDocumentKind kind, string path) {
        string[] visible = ["-34,50\u00a0zł", "12,50\u00a0USD", "(1\u00a0234\u00a0JPY)", "-7,25\u00a0USD"];
        if (kind == IWorkDocumentKind.Pages) {
            using var doc = WordprocessingDocument.Open(path, false);
            var rows = doc.MainDocumentPart!.Document!.Descendants<W.TableRow>().ToArray();
            Assert.Equal(5, rows.Length);
            Assert.Equal(visible, rows.Skip(1).Select(row => row.Elements<W.TableCell>().Last().InnerText));
        } else if (kind == IWorkDocumentKind.Keynote) {
            using var doc = PresentationDocument.Open(path, false);
            var rows = doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.TableRow>().ToArray();
            Assert.Equal(5, rows.Length);
            Assert.Equal(visible, rows.Skip(1).Select(row => row.Elements<A.TableCell>().Last().InnerText));
        } else {
            using var doc = SpreadsheetDocument.Open(path, false);
            WorkbookPart workbook = doc.WorkbookPart!;
            var cells = workbook.WorksheetParts.Single().Worksheet!.Descendants<S.Cell>()
                .Where(cell => cell.DataType == null || cell.DataType.Value == S.CellValues.Number)
                .Where(cell => cell.CellValue is not null).ToArray();
            Assert.Equal(new[] { "-34.5", "12.5", "-1234", "-7.25" }, cells.Select(cell => cell.CellValue!.Text));
            var styles = workbook.WorkbookStylesPart!.Stylesheet!;
            var codes = styles.NumberingFormats!.Elements<S.NumberingFormat>()
                .ToDictionary(format => format.NumberFormatId!.Value, format => format.FormatCode!.Value);
            Assert.Equal(new[] {
                "0.00 [$zł-415]", "[$$-409] 0.00", @"[$¥-411] #,##0_);\([$¥-411] #,##0\)", "[$$-409] 0.00"
            }, cells.Select(cell => codes[styles.CellFormats!.Elements<S.CellFormat>()
                .ElementAt((int)cell.StyleIndex!.Value).NumberFormatId!.Value]));
        }
    }

    private static byte[] Create(IWorkDocumentKind kind) {
        var data = IWorkTableData.Create(5, 2)
            .SetText(1, 1, "Currency").SetText(1, 2, "Amount")
            .SetText(2, 1, "PLN").SetNumber(2, 2, -12.5, IWorkNumberFormat.CreateCurrency("PLN", 2))
            .SetText(3, 1, "USD").SetNumber(3, 2, 12.5, IWorkNumberFormat.CreateCurrency("USD", 2))
            .SetText(4, 1, "JPY").SetNumber(4, 2, -1234, IWorkNumberFormat.CreateCurrency("JPY", 0, true, IWorkNegativeNumberStyle.Parentheses))
            .SetText(5, 1, "USD credit").SetNumber(5, 2, -7.25, IWorkNumberFormat.CreateCurrency("USD", 2));
        if (kind == IWorkDocumentKind.Numbers) {
            var document = IWorkNumbersDocument.Create();
            document.AddSheet("Amounts").AddTable(data, "Amounts", 20, 20, 180, 24);
            return document.ToBytes();
        }
        if (kind == IWorkDocumentKind.Pages) {
            var document = IWorkPagesDocument.Create();
            document.AddTable(data, "Amounts", 180, 24, fontSizePoints: 12);
            return document.ToBytes();
        }
        var deck = IWorkKeynoteDocument.Create();
        deck.AddSlide().AddTable(data, 120, 40, 600, 150, "Amounts", fontSizePoints: 18);
        return deck.ToBytes();
    }

    private static string Compact(string text) => string.Concat(text.Where(character => !char.IsWhiteSpace(character)));
    private static string Corpus(string name) => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", "native-exports", "currency-v16", name);
}
