using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Fact]
    public void LaterTableReferencePreservesEarlierSpanPageAndCoordinatesAgainstUnreferencedControl() {
        static byte[] Build(bool associateNote) {
            var document = PdfDocument.Create(Options());
            var note = document.Content.DefineFootnote(body => body.Text("AttachedNote", style: Style()));
            document.Content.Text("Host", style: Style()).Spacer(global::OfficeIMO.OfficeLength.Points(75D));
            var reference = PdfTextRun.Normal("LateReference");
            if (associateNote) reference = reference.WithFootnoteReference(note);
            var style = FootnoteTableStyle();
            style.MinimumBodyRowsOnFirstPage = 0;
            style.MinimumBodyRowsOnLastPage = 0;
            style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(110D), global::OfficeIMO.OfficeLength.Points(110D) };
            document.Content.Table(new[] {
                new[] {
                    new PdfTableCell(string.Join("\n", Enumerable.Range(1, 6).Select(index => "Prefix" + index.ToString("00"))), rowSpan: 2),
                    new PdfTableCell("RightFirst")
                },
                new[] { new PdfTableCell("RightLast") },
                new[] { new PdfTableCell(new[] { reference }, columnSpan: 2) }
            }, style: style);
            return document.ToBytes();
        }

        using var control = PdfPigDocument.Open(Build(associateNote: false));
        using var associated = PdfPigDocument.Open(Build(associateNote: true));
        Assert.Equal(2, control.NumberOfPages);
        Assert.Equal(control.NumberOfPages, associated.NumberOfPages);
        Assert.DoesNotContain("Prefix01", control.GetPage(1).Text);
        foreach (string marker in new[] { "Host", "RightFirst", "RightLast" }.Concat(Enumerable.Range(1, 6).Select(index => "Prefix" + index.ToString("00")))) {
            var controlPage = Assert.Single(control.GetPages(), page => page.GetWords().Any(word => word.Text == marker));
            var actualPage = Assert.Single(associated.GetPages(), page => page.GetWords().Any(word => word.Text == marker));
            Assert.Equal(controlPage.Number, actualPage.Number);
            var expected = Assert.Single(controlPage.GetWords(), word => word.Text == marker).Letters[0].StartBaseLine;
            var actual = Assert.Single(actualPage.GetWords(), word => word.Text == marker).Letters[0].StartBaseLine;
            Assert.Equal(expected.X, actual.X, 3);
            Assert.Equal(expected.Y, actual.Y, 3);
        }
        Assert.Contains("LateReference", associated.GetPage(2).Text);
        Assert.Contains("AttachedNote", associated.GetPage(2).Text);
        Assert.DoesNotContain("AttachedNote", control.GetPage(2).Text);
    }
}
