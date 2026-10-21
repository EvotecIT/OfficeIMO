using OfficeIMO.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    [Fact]
    public void SaveAsPdf_PageFootnoteReferenceRemainsBetweenItsSourceRuns() {
        var bytes = ConvertPageFootnoteFixture("footnote-doc-beneathText-section-pageBottom");
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        var page = Assert.Single(actual.GetPages());
        AssertPageFootnoteInlineReference(Assert.Single(reference.GetPages()), "REFERENCE PREFIX", "REFERENCE SUFFIX", "1");
        AssertPageFootnoteInlineReference(page, "REFERENCE PREFIX", "REFERENCE SUFFIX", "1");
        Assert.Contains("FIRST BODY", page.Text, StringComparison.Ordinal);
        Assert.Contains("FOLLOW BODY", page.Text, StringComparison.Ordinal);
        Assert.Contains("FOOTER", page.Text, StringComparison.Ordinal);
        AssertPageFootnoteNativeProfile(bytes);
    }

    [Theory]
    [InlineData("footnote-doc-beneathText-section-pageBottom")]
    [InlineData("footnote-doc-pageBottom-section-beneathText")]
    [InlineData("fallback-doc-beneathText-section-absent")]
    [InlineData("fallback-doc-beneathText-sectionFootnotePr-absent")]
    [InlineData("fallback-doc-absent-sectionFootnotePr-absent")]
    public void SaveAsPdf_PageFootnotePlacementUsesSectionPositionAndPageBottomDefault(string fixture) {
        // The conflicting document setting distinguishes section placement from
        // inheritance. Separate absent-position and absent-properties sources
        // protect the native page-bottom default for both reachable OOXML forms.
        var bytes = ConvertPageFootnoteFixture(fixture);
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        var page = Assert.Single(actual.GetPages());
        var native = Assert.Single(reference.GetPages());
        AssertPageFootnoteTokenOnce(actual, "NOTE BODY");
        AssertPageFootnotePortableStyle(page, "NOTE BODY");
        AssertPageFootnoteBodyDoesNotOverlapNotes(actual, new[] { "FIRST BODY", "REFERENCE PREFIX", "FOLLOW BODY" }, new[] { "NOTE BODY" });
        Assert.True(PageFootnoteToken(page, "NOTE BODY")[0].StartBaseLine.Y >
            PageFootnoteToken(page, "FOOTER")[0].StartBaseLine.Y);
        bool beneathText = fixture == "footnote-doc-pageBottom-section-beneathText";
        if (!beneathText) AssertPageFootnoteBottomLine(page, "NOTE BODY", bytes.Frame);
        var controlBytes = fixture == "footnote-doc-beneathText-section-pageBottom" ? bytes :
            ConvertPageFootnoteFixture("footnote-doc-beneathText-section-pageBottom");
        using var bottomControl = PdfPigDocument.Open(controlBytes.Actual);
        var expected = PageFootnoteToken(bottomControl.GetPage(1), "NOTE BODY")[0];
        var observed = PageFootnoteToken(page, "NOTE BODY")[0];
        if (beneathText) Assert.True(observed.StartBaseLine.Y > expected.StartBaseLine.Y);
        else {
            // The omitted forms must behave like explicit section pageBottom
            // with the same host fonts, regardless of native metric parity.
            Assert.Equal(expected.StartBaseLine.X, observed.StartBaseLine.X, 3);
            Assert.Equal(expected.StartBaseLine.Y, observed.StartBaseLine.Y, 3);
        }
        AssertPageFootnoteNativeProfile(bytes, () => {
            AssertPageFootnoteNativeText(page, native, "NOTE BODY",
                beneathText ? PageFootnoteNativeAnchor.PrecedingBodyBaseline : PageFootnoteNativeAnchor.PhysicalBottom,
                precedingBodyToken: beneathText ? "FOLLOW BODY" : null);
            AssertPageFootnoteNativeText(page, native, "FOOTER", PageFootnoteNativeAnchor.PhysicalBottom);
        });
    }

    [Fact]
    public void SaveAsPdf_PageFootnotesRetainReferenceOrderAndDistinctParagraphs() {
        var bytes = ConvertPageFootnoteFixture("richnote-two-notes-one-page");
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        var page = Assert.Single(actual.GetPages());
        var native = Assert.Single(reference.GetPages());
        AssertPageFootnoteInlineReference(page, "FIRST REFERENCE", "FIRST TAIL", "1");
        AssertPageFootnoteInlineReference(page, "SECOND REFERENCE", "SECOND TAIL", "2");
        string[] paragraphs = { "NOTE FIRST LINE", "NOTE FIRST LAST", "NOTE SECOND LINE" };
        foreach (string token in paragraphs) {
            AssertPageFootnoteTokenOnce(actual, token);
            AssertPageFootnotePortableStyle(page, token);
        }
        AssertPageFootnoteParagraphOrder(actual, paragraphs, bytes.Frame);
        AssertPageFootnoteBodyDoesNotOverlapNotes(actual,
            new[] { "BEFORE TWO", "FIRST REFERENCE", "SECOND REFERENCE", "FOLLOW TWO" }, paragraphs);
        Assert.Contains("FOLLOW TWO", page.Text, StringComparison.Ordinal);
        AssertPageFootnoteNativeProfile(bytes, () => {
            foreach (string token in paragraphs)
                AssertPageFootnoteNativeText(page, native, token, PageFootnoteNativeAnchor.PhysicalBottom);
        });
    }

    [Fact]
    public void SaveAsPdf_PageFootnoteRetainsItsOwnPictureAndInternalBacklink() {
        var bytes = ConvertPageFootnoteFixture("richnote-picture-internal-link");
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        var page = Assert.Single(actual.GetPages());
        var native = Assert.Single(reference.GetPages());
        var picture = Assert.Single(page.GetImages()).BoundingBox;
        var nativePicture = Assert.Single(native.GetImages()).BoundingBox;
        Assert.Equal(24D, picture.Width, 3);
        Assert.Equal(12D, picture.Height, 3);
        foreach (string token in new[] { "NOTE PICTURE", "AFTER IMAGE", "NOTE LINK", "BACK TO BODY" }) {
            AssertPageFootnoteTokenOnce(actual, token);
            AssertPageFootnotePortableStyle(page, token, color: token == "BACK TO BODY" ? 0x0000FF : 0);
        }
        Letter[] noteLead = PageFootnoteToken(page, "NOTE PICTURE");
        Letter[] noteTail = PageFootnoteToken(page, "AFTER IMAGE");
        Assert.True(picture.Left >= noteLead[noteLead.Length - 1].EndBaseLine.X);
        Assert.True(picture.Right <= noteTail[0].StartBaseLine.X);
        Assert.True(picture.Bottom <= noteLead[0].StartBaseLine.Y && picture.Top >= noteLead[0].StartBaseLine.Y);
        AssertPageFootnoteBottomLine(page, "BACK TO BODY", bytes.Frame);
        AssertPageFootnoteBodyDoesNotOverlapNotes(actual,
            new[] { "BODY ANCHOR", "PICTURE NOTE REFERENCE", "FOLLOW PICTURE NOTE" },
            new[] { "NOTE PICTURE", "AFTER IMAGE", "NOTE LINK", "BACK TO BODY" });
        Assert.True(PageFootnoteToken(page, "FOLLOW PICTURE NOTE").Min(letter => letter.BoundingBox.Bottom) >= picture.Top);
        var linkText = PageFootnoteToken(page, "BACK TO BODY");
        double linkLeft = linkText[0].StartBaseLine.X;
        double linkRight = linkText[linkText.Length - 1].EndBaseLine.X;
        double linkBaseline = linkText[0].StartBaseLine.Y;
        var underlines = page.Paths.Where(path =>
            ((path.IsFilled && path.FillColor.ToRGBValues() == (0D, 0D, 1D)) ||
             (path.IsStroked && path.StrokeColor.ToRGBValues() == (0D, 0D, 1D))))
            .Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue && bounds.Value.Top < linkBaseline &&
                bounds.Value.Bottom > linkBaseline - linkText[0].PointSize &&
                bounds.Value.Right >= linkLeft - .15D && bounds.Value.Left <= linkRight + .15D)
            .Select(bounds => bounds!.Value).OrderBy(bounds => bounds.Left).ToArray();
        Assert.NotEmpty(underlines);
        Assert.InRange(Math.Abs(underlines[0].Left - linkLeft), 0D, .15D);
        double underlineY = (underlines[0].Top + underlines[0].Bottom) / 2D;
        double coveredRight = underlines[0].Left;
        // Word and space segments may emit adjacent PDF paths. Their union
        // must cover the full actual link span continuously at one height.
        foreach (var bounds in underlines) {
            Assert.InRange(Math.Abs((bounds.Top + bounds.Bottom) / 2D - underlineY), 0D, .15D);
            Assert.True(bounds.Left <= coveredRight + .15D, "The link underline contains a visible gap.");
            coveredRight = Math.Max(coveredRight, bounds.Right);
        }
        Assert.InRange(Math.Abs(coveredRight - linkRight), 0D, .15D);
        PdfReadDocument read = PdfReadDocument.Open(bytes.Actual);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes.Actual);
        PdfLinkAnnotation backlink = Assert.Single(info.LinkAnnotations, link =>
            link.IsInternalDestinationLink &&
            link.X1 <= linkText[0].StartBaseLine.X + .15D &&
            link.X2 >= linkText[linkText.Length - 1].EndBaseLine.X - .15D &&
            link.Y1 <= linkText[0].StartBaseLine.Y + .15D && link.Y2 >= linkText[0].StartBaseLine.Y - .15D);
        PdfNamedDestination? namedTarget = backlink.DestinationName == null ? null :
            Assert.Single(read.NamedDestinations, destination => destination.Name == backlink.DestinationName);
        int? destinationPage = namedTarget?.PageNumber ?? backlink.DestinationPageNumber;
        double? destinationTop = namedTarget?.DestinationTop ?? backlink.DestinationTop;
        Assert.Equal(1, destinationPage);
        Assert.NotNull(destinationTop);
        double bodyBaseline = PageFootnoteToken(page, "BODY ANCHOR")[0].StartBaseLine.Y;
        Assert.InRange(destinationTop!.Value, bodyBaseline,
            bodyBaseline + PageFootnoteToken(page, "BODY ANCHOR")[0].PointSize);
        Assert.True(destinationTop.Value > linkText[0].StartBaseLine.Y);
        Assert.Equal(2, info.LinkAnnotations.Count);
        Assert.All(info.LinkAnnotations, link => Assert.True(link.IsInternalDestinationLink));
        Assert.Empty(read.Outlines);
        Assert.Empty(read.FormFields);
        AssertPageFootnoteNativeProfile(bytes, () => {
            Assert.InRange(Math.Abs(nativePicture.Left - picture.Left), 0D, .15D);
            Assert.InRange(Math.Abs(nativePicture.Top - picture.Top), 0D, .15D);
            foreach (string token in new[] { "NOTE PICTURE", "AFTER IMAGE", "NOTE LINK", "BACK TO BODY" })
                AssertPageFootnoteNativeText(page, native, token, PageFootnoteNativeAnchor.PhysicalBottom);
        });
    }
}
