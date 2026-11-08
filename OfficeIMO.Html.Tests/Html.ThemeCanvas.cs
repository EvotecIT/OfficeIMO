using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlThemeCanvasTests {
    [Fact]
    public void HtmlSnapshotsPreserveSolidCanvasAndTextColorsInNewWordDocuments() {
        HtmlConversionDocument dark = HtmlConversionDocument.Parse("<p>CANVAS</p>").WithTheme(OfficeTheme.Dark);
        var sources = new[] {
            (Source: dark, Canvas: "0F1620", Text: "E7ECF2"),
            (Source: HtmlConversionDocument.Parse("<style>body{background-color:rgba(32,48,64,.5);color:#102030}</style><p>CANVAS</p>"), Canvas: "9098A0", Text: "102030"),
            (Source: HtmlConversionDocument.Parse("<html style='--unused:1;background-color:#123456'><head><style>" +
                "@layer report{body{background-color:#ABCDEF;color:#FEDCBA}}</style></head><body><p>CANVAS</p></body></html>"), Canvas: "123456", Text: "FEDCBA")
        };
        foreach (var source in sources) {
            // Deliberately omit HtmlWordExportOptions.Theme: this converts the snapshot's actual CSS.
            using WordDocument document = source.Source.ToWord(new HtmlWordExportOptions {
                UnsupportedCssHandling = WordHtmlUnsupportedCssHandling.Error
            });
            byte[] bytes = document.ToBytes();
            using (WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(bytes), false)) {
                var main = package.MainDocumentPart!;
                Assert.Equal(source.Canvas, main.Document.DocumentBackground!.Color!.Value);
                W.Run run = Assert.Single(main.Document.Body!.Descendants<W.Run>(), run => run.InnerText == "CANVAS");
                Assert.Equal(source.Text, run.RunProperties!.Color!.Val!.Value);
            }
            using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
            using PdfCore.PdfDocument pdfModel = reopened.ToPdf(new WordPdfExportOptions {
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(), TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
            });
            using PdfPigDocument pdf = PdfPigDocument.Open(pdfModel.ToBytes());
            var page = pdf.GetPage(1);
            Assert.Equal("CANVAS", string.Concat(page.Letters.Select(letter => letter.Value)));
            Assert.Contains(page.Paths, path => path.IsFilled && path.FillColor is { } fill && RgbEquals(fill.ToRGBValues(), OfficeColor.ParseHex(source.Canvas)));
            Assert.All(page.Letters, letter => Assert.True(RgbEquals(letter.Color.ToRGBValues(), OfficeColor.ParseHex(source.Text))));
        }

        using WordDocument plain = HtmlConversionDocument.Parse("<p>Plain</p>").ToWord();
        Assert.Null(plain.Background.Color);
        using WordDocument host = WordDocument.Create();
        host.Background.Color = OfficeColor.Yellow;
        host.AddHtml(dark);
        Assert.Equal(OfficeColor.Yellow, host.Background.Color);
    }

    private static bool RgbEquals((double R, double G, double B) actual, OfficeColor expected) =>
        Math.Abs(actual.R - expected.R / 255D) < .001D && Math.Abs(actual.G - expected.G / 255D) < .001D &&
        Math.Abs(actual.B - expected.B / 255D) < .001D;
}
