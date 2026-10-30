using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Fact]
    public void ImportedPictureClipUsesTheCurrentParagraphFragmentTopAfterNoteContinuation() {
        var document = PdfDocument.Create(Options());
        var picture = new PdfInlineImage(PdfPngTestImages.CreateRgbPng((byte)0, (byte)180, (byte)80), 16, 48) {
            AllowFrameOverflow = true
        };
        PdfTextRun[] runs = Enumerable.Range(0, 10).SelectMany(index => index is 1 or 8
            ? new[] { PdfTextRun.Normal($"BEFORE{index} "), PdfTextRun.Inline(picture), PdfTextRun.Normal($" AFTER{index}\n") }
            : new[] { PdfTextRun.Normal($"PLAIN{index}" + (index < 9 ? "\n" : "")) }).ToArray();
        var note = document.Content.DefineFootnote(body => body.Paragraph(paragraph => paragraph.Runs(runs),
            style: new PdfParagraphStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(10), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(12)), WidowControl = false,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            }));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        byte[] bytes = document.ToBytes();
        PdfImagePlacement[] images = PdfImageExtractor.ExtractImagePlacements(bytes).OrderBy(image => image.PageNumber).ToArray();
        Assert.Equal(2, images.Length);
        Assert.Equal(new[] { 1, 2 }, images.Select(image => image.PageNumber));
        PdfImageClipInfo first = Assert.IsType<PdfImageClipInfo>(images[0].Clip);
        PdfImageClipInfo continued = Assert.IsType<PdfImageClipInfo>(images[1].Clip);
        Assert.True(first.IsExact && first.IsRectangle && continued.IsExact && continued.IsRectangle);
        // Both pictures are on the second line of their physical paragraph
        // fragment. A per-line top clips too much; a source-global top on page2
        // allows paint into the preceding page-note/separator area.
        Assert.InRange(first.Height, 12.01D, 24.01D);
        Assert.Equal(first.Height, continued.Height, 3);
        Assert.Equal(16D, first.Width, 3);
        Assert.Equal(16D, continued.Width, 3);
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("PLAIN0", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("PLAIN7", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("PLAIN7", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("AFTER8", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("PLAIN9", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }
}
