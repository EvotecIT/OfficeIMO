using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void PdfTables_SaveAsWord_ImportsDetectedTablesAsWordTables() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        using var document = new MemoryStream();
        IReadOnlyList<PdfWordTableImportResult> results = PdfWordTableConverterExtensions.SaveAsWordFromPdfTables(
            pdf,
            document,
            new PdfWordTableImportOptions {
                LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            });

        PdfWordTableImportResult result = Assert.Single(results);
        Assert.Equal(1, result.PageNumber);
        Assert.Equal(0, result.TableIndex);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(2, result.RowCount);
        Assert.False(result.Truncated);
        Assert.True(result.HeaderRowIncluded);

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(document.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());

        Body body = GetBody(package);
        Table table = Assert.Single(body.Descendants<Table>());
        List<TableRow> rows = table.Elements<TableRow>().ToList();
        Assert.Equal(3, rows.Count);
        Assert.NotNull(rows[0].TableRowProperties?.GetFirstChild<TableHeader>());
        Assert.Equal(new[] { "Code", "Name", "Qty" }, ReadRowText(rows[0]));
        Assert.Equal(new[] { "A-100", "Alpha", "2" }, ReadRowText(rows[1]));
        Assert.Equal(new[] { "B-200", "Beta", "14" }, ReadRowText(rows[2]));
        Assert.Null(ReadCellAlignment(rows[0], 2));
        Assert.Null(ReadCellAlignment(rows[1], 1));
        Assert.Equal(JustificationValues.Right, ReadCellAlignment(rows[1], 2));
        Assert.Equal(JustificationValues.Right, ReadCellAlignment(rows[2], 2));
        Assert.Contains(body.Descendants<Text>(), text => text.Text == "PDF page 1, table 1");
    }

    [Fact]
    public void PdfTables_SaveAsWord_AppliesRowCapsAndKeepsDocumentValidWhenEmpty() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfCore.PdfKeyValueRow.Text("Due", "2026-06-30")
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 170 },
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .PageBreak()
            .Paragraph(p => p.Text("No table on this page."))
            .ToBytes();

        using var document = new MemoryStream();
        IReadOnlyList<PdfWordTableImportResult> results = PdfWordTableConverterExtensions.SaveAsWordFromPdfTables(
            pdf,
            document,
            new PdfWordTableImportOptions {
                LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                    ForceSingleColumn = true
                },
                PageRanges = new[] { PdfCore.PdfPageRange.From(1, 1) },
                MaxRows = 2,
                IncludeSourceCaptions = false
            });

        PdfWordTableImportResult result = Assert.Single(results);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(3, result.TotalRowCount);
        Assert.True(result.Truncated);

        using (WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(document.ToArray()), false)) {
            Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
            Table table = Assert.Single(GetBody(package).Descendants<Table>());
            List<TableRow> rows = table.Elements<TableRow>().ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal(new[] { "Key", "Value" }, ReadRowText(rows[0]));
            Assert.Equal(new[] { "InvoiceId", "INV-001" }, ReadRowText(rows[1]));
            Assert.Equal(new[] { "Customer", "Evotec" }, ReadRowText(rows[2]));
        }

        using var emptyDocument = new MemoryStream();
        IReadOnlyList<PdfWordTableImportResult> emptyResults = PdfWordTableConverterExtensions.SaveAsWordFromPdfTables(
            pdf,
            emptyDocument,
            new PdfWordTableImportOptions {
                LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                    ForceSingleColumn = true
                },
                PageRanges = new[] { PdfCore.PdfPageRange.From(2, 2) },
                EmptyDocumentMessage = "Nothing tabular was detected."
            });

        Assert.Empty(emptyResults);
        using WordprocessingDocument emptyPackage = WordprocessingDocument.Open(new MemoryStream(emptyDocument.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(emptyPackage).ToList());
        Body emptyBody = GetBody(emptyPackage);
        Assert.Empty(emptyBody.Descendants<Table>());
        Assert.Contains(emptyBody.Descendants<Text>(), text => text.Text == "Nothing tabular was detected.");
    }

    private static Body GetBody(WordprocessingDocument package) {
        return package.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("The saved document body is missing.");
    }

    private static string[] ReadRowText(TableRow row) {
        return row.Elements<TableCell>()
            .Select(cell => string.Concat(cell.Descendants<Text>().Select(text => text.Text ?? string.Empty)))
            .ToArray();
    }

    private static JustificationValues? ReadCellAlignment(TableRow row, int columnIndex) {
        return row.Elements<TableCell>()
            .ElementAt(columnIndex)
            .Elements<Paragraph>()
            .FirstOrDefault()?
            .ParagraphProperties?
            .Justification?
            .Val?
            .Value;
    }
}
