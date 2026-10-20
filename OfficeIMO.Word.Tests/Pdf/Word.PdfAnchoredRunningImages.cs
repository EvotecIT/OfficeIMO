using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("header-anchor-explicit-defaults.docx", 21.45D, 40.20D, 66.36D)]
    [InlineData("header-anchor-explicit-defaults-preceding-line.docx", 33.45D, 52.20D, 67.32D)]
    public void SaveAsPdf_RunningHeaderPictureRetainsNativeParagraphAnchorAndRightTab(
        string fixture, double imageTop, double titleBaseline, double bodyBaseline) {
        // Fully synthetic sources explicitly declare paragraph defaults and
        // were schema-validated, exported by native Word, and inspected
        // independently. A preceding header paragraph moves
        // the paragraph-relative picture and title together, while body entry
        // follows the header's text rather than the picture's 49pt paint frame.
        string path = Path.Combine(AppContext.BaseDirectory, "WordLayoutFixtures", fixture);
        byte[] source = File.ReadAllBytes(path);
        using WordDocument document = WordDocument.Load(path);
        Assert.Empty(document.ValidateDocument());
        string sourceXml = AnchoredRunningImageSourceXml(document);

        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var picture = Assert.Single(page.GetImages());
        Assert.InRange(Math.Abs(picture.BoundingBox.Left - 28.35D), 0D, .15D);
        Assert.InRange(Math.Abs(page.Height - picture.BoundingBox.Top - imageTop), 0D, .15D);
        Assert.Equal(128D, picture.BoundingBox.Width, 3);
        Assert.Equal(49D, picture.BoundingBox.Height, 3);

        var words = page.GetWords().ToArray();
        var title = Assert.Single(words, word => word.Text == "PROJECT");
        var titleEnd = Assert.Single(words, word => word.Text == "INITIATIVE");
        var body = Assert.Single(words, word => word.Text == "BODY");
        Assert.InRange(Math.Abs(title.Letters[0].StartBaseLine.X - 416.35D), 0D, .3D);
        Assert.InRange(Math.Abs(titleEnd.BoundingBox.Right - 581.2D), 0D, .3D);
        Assert.InRange(Math.Abs(page.Height - title.Letters[0].StartBaseLine.Y - titleBaseline), 0D, .3D);
        Assert.InRange(Math.Abs(page.Height - body.Letters[0].StartBaseLine.Y - bodyBaseline), 0D, .3D);
        Assert.InRange(Math.Abs(body.Letters[0].StartBaseLine.X - 28.35D), 0D, .15D);
        Assert.Contains("Generic form content.", page.Text);
        if (fixture.EndsWith("preceding-line.docx", StringComparison.Ordinal)) {
            var preceding = Assert.Single(words, word => word.Text == "HEADER");
            Assert.InRange(Math.Abs(page.Height - preceding.Letters[0].StartBaseLine.Y - 30.84D), 0D, .3D);
        }
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeAnchoredImageFlowed");
        Assert.Equal(sourceXml, AnchoredRunningImageSourceXml(document));
        Assert.Equal(source, File.ReadAllBytes(path));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_UnqualifiedRunningPictureKeepsBoundedImageAndTextProjection(bool columnDependent) {
        using WordDocument source = WordDocument.Load(Path.Combine(AppContext.BaseDirectory,
            "WordLayoutFixtures", "header-anchor-explicit-defaults.docx"));
        DW.Anchor anchor = source._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single()
            .Header.Descendants<DW.Anchor>().Single();
        if (columnDependent) {
            source.Sections[0].ColumnCount = 2;
        } else {
            anchor.RemoveAllChildren<DW.WrapNone>();
            anchor.AddChild(new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides }, true);
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        string sourceXml = AnchoredRunningImageSourceXml(document);
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var page = Assert.Single(pdf.GetPages());
        var picture = Assert.Single(page.GetImages());
        // Wrapping and a not-yet-selected text column are outside the fixed
        // running-picture contract. Preserve the existing readable projection
        // instead of claiming the qualified left-margin position for them.
        Assert.InRange(picture.BoundingBox.Left, 160D, page.Width - picture.BoundingBox.Width);
        foreach (string token in new[] { "PROJECT", "INITIATIVE", "BODY", "ANCHOR" })
            Assert.Single(page.GetWords(), word => word.Text == token);
        Assert.Equal(sourceXml, AnchoredRunningImageSourceXml(document));
    }

    private static string AnchoredRunningImageSourceXml(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        return main.Document.OuterXml + main.DocumentSettingsPart!.Settings!.OuterXml +
            string.Concat(main.HeaderParts.Select(part => part.Header.OuterXml));
    }
}
