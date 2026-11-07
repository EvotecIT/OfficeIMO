using OfficeIMO.Pdf.Word;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Word.Pdf;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalTaggedTableBodyProjectionTests {
    [Fact]
    public void NativeTaggedTablesPreserveRowsWithPaintedEmptyCells() {
        string fixture = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents",
            "IWorkCorpus", "native-exports", "numbers-created-v14.5.pdf");
        using PdfDocument document = PdfDocument.Load(fixture);
        PdfDocumentReadResult read = document.Read();

        Assert.Collection(read.Pages[0].Tables,
            items => {
                Assert.Collection(items.Rows,
                    row => Assert.Equal(new[] { "Item", "Value", "Flag", "Notes" }, row),
                    row => Assert.Equal(new[] { "Apples", "20,75", "PRAWDA", "Zażółć 😀" }, row),
                    row => Assert.Equal(new[] { "Negative", "-0,125", "FAŁSZ", "=literal" }, row),
                    row => Assert.Equal(new[] { "Line1 Line2", "12345,6789", "PRAWDA", "" }, row));
            },
            totals => Assert.Collection(totals.Rows,
                row => Assert.Equal(new[] { "Totals", "99" }, row),
                row => Assert.Equal(new[] { "FAŁSZ", "" }, row)));
        Assert.Contains("Negative\t-0,125\tFAŁSZ\t=literal", read.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Pusta komórka", read.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void NativeTaggedTableSupersedesGeometryThatAbsorbsItsCaption() {
        string fixture = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents",
            "IWorkCorpus", "native-exports", "numbers-formulas-v14.5.pdf");
        using PdfDocument document = PdfDocument.Load(fixture);
        PdfDocumentReadResult read = document.Read();

        var table = Assert.Single(read.Pages[0].Tables);
        Assert.Equal(new[] { "3", "2", "30.09.2026 12:42" }, table.Rows[2]);
        Assert.Contains("Table 1" + Environment.NewLine + "1\t1\tPRAWDA", read.Text, StringComparison.Ordinal);
        Assert.Single(Regex.Matches(read.Text, Regex.Escape("30.09.2026 12:42")).Cast<Match>());
    }

    [Fact]
    public void NativeTaggedTablesPreserveClearedPublisherCells() {
        string fixture = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Publisher.Tests", "Fixtures",
            "NativeComparison", "TableCellTypingEarlyRange", "native-tablecelltyping.pdf");
        using PdfDocument document = PdfDocument.Load(fixture);
        PdfDocumentReadResult read = document.Read();

        Assert.Collection(read.Pages[0].Tables,
            baseline => Assert.Equal(new[] { "Item", "Count" }, baseline.Rows[0]),
            clearFirst => {
                Assert.Equal(new[] { "", "Count" }, clearFirst.Rows[0]);
                Assert.Equal(new[] { "Alpha", "2" }, clearFirst.Rows[1]);
            },
            clearFinal => {
                Assert.Equal(new[] { "Item", "Count" }, clearFinal.Rows[0]);
                Assert.Equal(new[] { "Alpha", "" }, clearFinal.Rows[1]);
            },
            styledFirst => Assert.Equal(new[] { "", "Count" }, styledFirst.Rows[0]));
        Assert.Contains("ClearFirst" + Environment.NewLine + "\tCount", read.Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(20D, false)]
    [InlineData(20D, true)]
    [InlineData(150D, false)]
    [InlineData(150D, true)]
    public void TaggedTableBesideListBodyPreservesProseAndEmitsEachCellOnce(double bodyX, bool irregularColumns) {
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false }.EnableTaggedPdfCatalogMarkers())
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.List, list => list
                .Structure(PdfCanvasStructureRole.ListItem, item => item
                    .Structure(PdfCanvasStructureRole.ListLabel, label => label.Text("1.", global::OfficeIMO.OfficeLength.Points(bodyX), global::OfficeIMO.OfficeLength.Points(90D), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(16D)))
                    .Structure(PdfCanvasStructureRole.ListBody, body => body
                        .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph.Text("Unique body", global::OfficeIMO.OfficeLength.Points(bodyX), global::OfficeIMO.OfficeLength.Points(110D), global::OfficeIMO.OfficeLength.Points(100D), global::OfficeIMO.OfficeLength.Points(16D)))
                        .Structure(PdfCanvasStructureRole.Table, table => table
                            .Structure(PdfCanvasStructureRole.TableRow, row => row
                                .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Metric", global::OfficeIMO.OfficeLength.Points(250D), global::OfficeIMO.OfficeLength.Points(110D), global::OfficeIMO.OfficeLength.Points(80D), global::OfficeIMO.OfficeLength.Points(16D)))
                                .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Value", global::OfficeIMO.OfficeLength.Points(360D), global::OfficeIMO.OfficeLength.Points(110D), global::OfficeIMO.OfficeLength.Points(80D), global::OfficeIMO.OfficeLength.Points(16D))))
                            .Structure(PdfCanvasStructureRole.TableRow, row => row
                                .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("Quality", global::OfficeIMO.OfficeLength.Points(irregularColumns ? 280D : 250D), global::OfficeIMO.OfficeLength.Points(130D), global::OfficeIMO.OfficeLength.Points(80D), global::OfficeIMO.OfficeLength.Points(16D)))
                                .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("High", global::OfficeIMO.OfficeLength.Points(irregularColumns ? 330D : 360D), global::OfficeIMO.OfficeLength.Points(130D), global::OfficeIMO.OfficeLength.Points(80D), global::OfficeIMO.OfficeLength.Points(16D)))))))))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(new System.IO.MemoryStream(bytes));
        var options = new PdfLoadOptions { Profile = PdfReadProfile.Structured };
        PdfDocumentReadResult read = document.Read(options);
        using OfficeIMO.Word.WordDocument word = (read.ConvertToWordWithEvidence()).Value;
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);
        string[] projections = {
            read.Text,
            read.ExtractMarkdown(),
            read.ToHtml(new PdfHtmlExportOptions { Profile = PdfHtmlProfile.Semantic }),
            package.MainDocumentPart!.Document!.InnerText,
            document.Reader.Text()
        };
        Assert.All(projections, text => {
            foreach (string expected in new[] { "Unique body", "Metric", "Value", "Quality", "High" }) {
                Assert.Single(Regex.Matches(text, Regex.Escape(expected)).Cast<Match>());
            }
        });
        Assert.Single(read.Pages[0].Tables);
        var reader = PdfReaderAdapter.ReadDocument(read);
        string readerBody = string.Join(" ", reader.Blocks.Select(static block => block.Text));
        Assert.Single(Regex.Matches(readerBody, "Unique body").Cast<Match>());
        Assert.DoesNotContain("Metric", readerBody, StringComparison.Ordinal);
        Assert.Single(reader.Tables);
    }
}
