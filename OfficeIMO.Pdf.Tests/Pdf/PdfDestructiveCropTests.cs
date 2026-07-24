using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDestructiveCropTests {
    [Fact]
    public void DestructiveCrop_ReplacesSelectedPageContentAndPreservesOtherStructures() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Selected page source marker"))
            .PageBreak()
            .Paragraph(p => p.Text("Untouched page marker"))
            .AttachFile("proof.txt", Encoding.UTF8.GetBytes("attachment proof"), "text/plain")
            .ToBytes();

        PdfDestructiveCropResult result = PdfDocument.Open(source).Pages.DestructiveCrop(
            0, 350, 612, 792,
            new PdfDestructiveCropOptions { Dpi = 96 },
            1);

        byte[] output = result.ToBytes();
        PdfReadDocument read = PdfReadDocument.Open(output);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.PageTreePlan.ExecutionMode);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.ContentPlan.ExecutionMode);
        Assert.True(result.PreservationReport.IsPreserved, string.Join(" ", result.PreservationReport.Issues.Select(static issue => issue.Message)));
        Assert.Single(result.Renders);
        Assert.Equal((612D, 442D), read.Pages[0].GetPageSize());
        Assert.Equal(string.Empty, read.Pages[0].ExtractText());
        Assert.NotEmpty(read.Pages[0].GetImages(1));
        Assert.Contains("Untouched page marker", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.Single(read.ExtractAttachments());
        Assert.DoesNotContain("Selected page source marker", read.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void DestructiveCrop_BlocksSelectedPagesWithFormWidgets() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Form crop"))
            .TextField("Account", width: 120, height: 20)
            .ToBytes();

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() => PdfPageEditor.DestructiveCropPages(source, 0, 0, 300, 500));
        Assert.Contains(exception.Plan.BlockerCodes, static code => code.Contains("Forms", StringComparison.Ordinal));
    }

    [Fact]
    public void DestructiveCrop_RemovesPageThumbnailReferenceAndPayload() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Thumbnail crop source")).ToBytes();
        int pageObjectNumber = PdfReadDocument.Open(source).Pages[0].ObjectNumber;
        byte[] withThumbnail = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            int thumbnailObjectNumber = objects.Keys.Max() + 1;
            var imageDictionary = new PdfDictionary();
            imageDictionary.Items["Type"] = new PdfName("XObject");
            imageDictionary.Items["Subtype"] = new PdfName("Image");
            imageDictionary.Items["Width"] = new PdfNumber(1);
            imageDictionary.Items["Height"] = new PdfNumber(1);
            imageDictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
            imageDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
            objects[thumbnailObjectNumber] = new PdfIndirectObject(
                thumbnailObjectNumber,
                0,
                new PdfStream(imageDictionary, new byte[] { 0x7F }));
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects[pageObjectNumber].Value);
            page.Items["Thumb"] = new PdfReference(thumbnailObjectNumber, 0);
            return security.InfoObjectNumber;
        });

        byte[] cropped = PdfPageEditor.DestructiveCropPages(withThumbnail, 0, 0, 612, 792, new PdfDestructiveCropOptions { Dpi = 72 }).ToBytes();

        Assert.DoesNotContain("/Thumb", PdfEncoding.Latin1GetString(cropped), StringComparison.Ordinal);
    }
}
