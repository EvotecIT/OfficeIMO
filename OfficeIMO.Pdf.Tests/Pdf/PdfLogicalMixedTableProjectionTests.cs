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
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }.EnableTaggedPdfCatalogMarkers())
            .Canvas(canvas => canvas
                .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph.Text("Unique body", bodyX, 110D, 100D, 16D))
                .Structure(PdfCanvasStructureRole.Table, table => table
                    .Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Metric", 250D, 110D, 80D, 16D))
                        .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Value", 360D, 110D, 80D, 16D)))
                    .Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("Quality", irregularColumns ? 280D : 250D, 130D, 80D, 16D))
                        .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("High", irregularColumns ? 330D : 360D, 130D, 80D, 16D)))))
    .ToBytes();
        PdfDocument document = PdfDocument.Load(bytes);
        var options = new PdfReadOptions { Profile = PdfReadProfile.Structured };
        PdfDocumentReadResult read = document.Read(options);
        using OfficeIMO.Word.WordDocument word = read.ToWordDocument();
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);
        string[] projections = {
            read.Text,
            read.ToMarkdown(),
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
