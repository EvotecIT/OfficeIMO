using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("table-legacy100-grid479-m575.docx", 479.5D, 335.65D, 72D)]
    [InlineData("table-legacy100-grid320-m575.docx", 479.5D, 335.65D, 72D)]
    [InlineData("table-legacy100-grid479-m0.docx", 468D, 327.6D, 72D)]
    [InlineData("table-legacy50-grid479-m575.docx", 239.75D, 167.8D, 72D)]
    [InlineData("table-modern100-grid479-m575.docx", 468D, 327.6D, 72D)]
    [InlineData("table-legacy100-grid479-m575-fixed.docx", 479.5D, 335.65D, 72D)]
    [InlineData("table-legacy50-grid479-m10.docx", 244D, 170.8D, 72D)]
    [InlineData("table-legacy100-grid479-m575-omitted-jc.docx", 479.5D, 335.65D, 72D)]
    [InlineData("table-legacy100-grid479-m575-indent-direct-zero.docx", 479.5D, 335.65D, 66.25D)]
    [InlineData("table-legacy100-grid479-m575-indent-defaultstyle-zero.docx", 479.5D, 335.65D, 66.25D)]
    [InlineData("table-indent-direct-six.docx", 479.5D, 335.65D, 72.25D)]
    [InlineData("table-indent-defaultstyle-six.docx", 479.5D, 335.65D, 66.25D)]
    [InlineData("table-normal-default-no-indent.docx", 479.5D, 335.65D, 66.25D)]
    [InlineData("table-custom-default-six.docx", 479.5D, 335.65D, 72.25D)]
    [InlineData("table-margin-table-asymmetric.docx", 486D, 340.2D, 72D)]
    [InlineData("table-margin-cell-override.docx", 486D, 340.2D, 72D)]
    public void SaveAsPdf_PercentageTableFramesMatchNativeWordAcrossBodyHeaderAndFooter(
        string fixture, double frameWidth, double firstCellWidth, double frameLeft) {
        // These saved DOCX controls were independently opened and exported by
        // native Word. Cache-only and compatibility variants distinguish the
        // percentage reference frame from stored grid widths and cell padding.
        string path = Path.Combine(AppContext.BaseDirectory, "WordLayoutFixtures", fixture);
        byte[] source = File.ReadAllBytes(path);
        using (var sourcePackage = WordprocessingDocument.Open(new MemoryStream(source), false)) {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(sourcePackage));
        }
        using WordDocument document = WordDocument.Load(path);
        string sourceXml = PercentageTableFrameSourceXml(document);

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Page page = Assert.Single(pdf.GetPages());
        AssertPercentageTableFrameGeometry(page, frameWidth, firstCellWidth, frameLeft);
        Assert.Equal(sourceXml, PercentageTableFrameSourceXml(document));
        Assert.Equal(source, File.ReadAllBytes(path));
    }

    private static void AssertPercentageTableFrameGeometry(Page page,
        double frameWidth, double firstCellWidth, double frameLeft) {
        PdfRectangle[] firstCells = PercentageTableFrameFills(page, first: true);
        PdfRectangle[] lastCells = PercentageTableFrameFills(page, first: false);
        Assert.Equal(3, firstCells.Length);
        Assert.Equal(3, lastCells.Length);
        for (int story = 0; story < firstCells.Length; story++) {
            PdfRectangle first = firstCells[story], last = lastCells[story];
            // Native Word stores table columns in twips; PDF paint coordinates
            // additionally reflect its export precision. This tolerance is well
            // below the observed 3.45-11.5 pt regression.
            Assert.InRange(Math.Abs(first.Left - frameLeft), 0D, .15D);
            Assert.InRange(Math.Abs(first.Width - firstCellWidth), 0D, .15D);
            Assert.InRange(Math.Abs(last.Width - (frameWidth - firstCellWidth)), 0D, .15D);
            Assert.InRange(Math.Abs(last.Right - first.Left - frameWidth), 0D, .15D);
            Assert.InRange(Math.Abs(last.Left - first.Right), 0D, .01D);
            Assert.InRange(Math.Abs(last.Top - first.Top), 0D, .01D);
        }

        string text = string.Concat(page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Select(letter => letter.Value));
        foreach (string label in new[] { "HEADERA", "HEADERB", "BODYA", "BODYB", "FOOTERA", "FOOTERB" })
            Assert.Contains(label, text);
        Assert.Contains("BEFOREBODY", text);
        Assert.Contains("AFTERBODY", text);
    }

    private static PdfRectangle[] PercentageTableFrameFills(Page page, bool first) =>
        page.Paths.Where(path => {
            if (!path.IsFilled || path.GetBoundingRectangle() is not { Width: > 20D, Height: > 1D }) return false;
            var (red, green, blue) = path.FillColor.ToRGBValues();
            return first
                ? red > .99D && Math.Abs(green - 128D / 255D) < .002D && Math.Abs(blue - 128D / 255D) < .002D
                : Math.Abs(red - 160D / 255D) < .002D && Math.Abs(green - 208D / 255D) < .002D && blue > .99D;
        }).Select(path => path.GetBoundingRectangle()!.Value).OrderByDescending(rectangle => rectangle.Top).ToArray();

    private static string PercentageTableFrameSourceXml(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        return main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml +
            main.DocumentSettingsPart!.Settings!.OuterXml +
            string.Concat(main.HeaderParts.Select(part => part.Header.OuterXml)) +
            string.Concat(main.FooterParts.Select(part => part.Footer.OuterXml));
    }
}
