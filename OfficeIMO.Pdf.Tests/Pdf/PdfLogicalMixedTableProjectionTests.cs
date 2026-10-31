using OfficeIMO.Pdf.Word;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalMixedTableProjectionTests {
    [Theory]
    [InlineData(20D, false)]
    [InlineData(20D, true)]
    [InlineData(150D, false)]
    [InlineData(150D, true)]
    public void TaggedTableBesideParagraphPreservesProseAndEmitsEachCellOnce(double bodyX, bool irregularColumns) {
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false }.EnableTaggedPdfCatalogMarkers())
            .Canvas(canvas => canvas
                .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph.Text("Unique body", global::OfficeIMO.OfficeLength.Points(bodyX), global::OfficeIMO.OfficeLength.Points(110D), global::OfficeIMO.OfficeLength.Points(100D), global::OfficeIMO.OfficeLength.Points(16D)))
                .Structure(PdfCanvasStructureRole.Table, table => table
                    .Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Metric", global::OfficeIMO.OfficeLength.Points(250D), global::OfficeIMO.OfficeLength.Points(110D), global::OfficeIMO.OfficeLength.Points(80D), global::OfficeIMO.OfficeLength.Points(16D)))
                        .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Value", global::OfficeIMO.OfficeLength.Points(360D), global::OfficeIMO.OfficeLength.Points(110D), global::OfficeIMO.OfficeLength.Points(80D), global::OfficeIMO.OfficeLength.Points(16D))))
                    .Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("Quality", global::OfficeIMO.OfficeLength.Points(irregularColumns ? 280D : 250D), global::OfficeIMO.OfficeLength.Points(130D), global::OfficeIMO.OfficeLength.Points(80D), global::OfficeIMO.OfficeLength.Points(16D)))
                        .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("High", global::OfficeIMO.OfficeLength.Points(irregularColumns ? 330D : 360D), global::OfficeIMO.OfficeLength.Points(130D), global::OfficeIMO.OfficeLength.Points(80D), global::OfficeIMO.OfficeLength.Points(16D))))))
    .ToBytes();
        PdfDocument document = PdfDocument.Load(new System.IO.MemoryStream(bytes));
        var options = new PdfLoadOptions { Profile = PdfReadProfile.Structured };
        PdfDocumentReadResult read = document.Read(options);
        using OfficeIMO.Word.WordDocument word = (read.ConvertToWordWithEvidence()).Value;
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);
        string[] projections = {
            read.Text,
            read.ExtractMarkdown(),
            read.ToHtml(new PdfToHtmlOptions { Profile = PdfHtmlProfile.Semantic }),
            package.MainDocumentPart!.Document!.InnerText,
            document.Reader.Text()
        };
        Assert.All(projections, text => {
            foreach (string expected in new[] { "Unique body", "Metric", "Value", "Quality", "High" }) {
                Assert.Single(Regex.Matches(text, Regex.Escape(expected)).Cast<Match>());
            }
        });
        Assert.Single(read.Pages[0].Tables);
    }
}
