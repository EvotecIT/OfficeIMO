using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfMirroredMarginResourceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MirroredContinuationReusesNamedFontsAndRetainsEveryPagesGlyphs(bool nested) {
        string fontPath = Assert.IsType<string>(PdfComplianceTestFonts.FindBundledTrueTypeFont());
        var options = Options().RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Mirror Font", File.ReadAllBytes(fontPath)));
        byte[] bytes = BuildDocument(options, 12, nested, namedFont: true);
        var fonts = PdfDocument.Load(bytes).Resources.Fonts();
        Assert.InRange(fonts.Fonts.Count, 1, 2);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(12, pdf.NumberOfPages);
        for (int page = 1; page <= 12; page++) Assert.Contains("Marker" + page, pdf.GetPage(page).Text);
    }

#if NET8_0_OR_GREATER && PDF_PERFORMANCE_EVIDENCE
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    [Trait("Category", "ResourcePerformanceEvidence")]
    public void MirroredContinuationDoesNotCopyDocumentAttachmentsForEveryPage(bool nested) {
        const int payloadLength = 2 * 1024 * 1024;
        byte[] payload = new byte[payloadLength];
        payload[0] = 42; payload[payload.Length - 1] = 27;
        var options = Options().AddEmbeddedFile("proof.bin", payload, "application/octet-stream");
        BuildDocument(options, 2, nested, namedFont: false);
        BuildDocument(options, 12, nested, namedFont: false);
        long shortRun = PdfAllocationTestSupport.MeasureMinimumThreadAllocation(() => BuildDocument(options, 2, nested, namedFont: false));
        long longRun = PdfAllocationTestSupport.MeasureMinimumThreadAllocation(() => BuildDocument(options, 12, nested, namedFont: false));
        Assert.True(longRun - shortRun < payloadLength * 2L,
            "Ten additional pages allocated " + (longRun - shortRun) + " bytes; document attachments must not be copied per page.");
        var attachment = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(BuildDocument(options, 12, nested, namedFont: false)));
        Assert.Equal(payload, attachment.Bytes);
    }
#endif

    private static PdfOptions Options() => new() {
        PageWidth = 300, PageHeight = 200,
        MarginLeft = 40, MarginRight = 70, MarginTop = 20, MarginBottom = 20,
        MirrorMargins = true, DefaultFontSize = 10
    };

    private static byte[] BuildDocument(PdfOptions options, int pages, bool nested, bool namedFont) {
        void AddPages(PdfContentBuilder content) {
            for (int page = 1; page <= pages; page++) {
                if (page > 1) content.PageBreak();
                content.Paragraph(paragraph => {
                    if (namedFont) paragraph.FontFamily("Mirror Font");
                    paragraph.Text("Marker" + page);
                });
            }
        }
        return PdfDocument.Create(builder => builder.Content(content => {
            if (nested) content.Element(element => element.Padding(2, 5).Content(AddPages));
            else AddPages(content);
        }), options).ToBytes();
    }
}
