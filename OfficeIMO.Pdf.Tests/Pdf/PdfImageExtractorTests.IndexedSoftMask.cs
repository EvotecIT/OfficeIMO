using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfImageExtractorTests {
    [Fact]
    public void LogicalReadResolvesRepeatedIndexedSoftMasksWithEachPagesPaletteAndMetadata() {
        using var input = new MemoryStream(BuildRepeatedIndexedSoftMaskPdf(), writable: false);
        using PdfDocument document = PdfDocument.Load(input);
        PdfDocumentReadResult read = document.Read();

        Assert.Equal(3, read.Pages.Count);
        PdfLogicalImage first = Assert.Single(read.Pages[0].Images);
        PdfLogicalImage second = Assert.Single(read.Pages[1].Images);
        PdfLogicalImage remapped = Assert.Single(read.Pages[2].Images);
        Assert.All(new[] { first, second, remapped }, image => {
            Assert.Equal(9, image.SourceImage.ObjectNumber);
            Assert.Equal("image/png", image.MimeType);
            Assert.True(image.SourceImage.TransparencyMaskResolved);
            Assert.Single(image.Placements);
        });
        Assert.Equal((1, "ImFirst", 36D), (first.PageNumber, first.ResourceName, first.PlacedX!.Value));
        Assert.Equal((2, "ImSecond", 72D), (second.PageNumber, second.ResourceName, second.PlacedX!.Value));
        Assert.Equal((3, "ImRemapped", 108D), (remapped.PageNumber, remapped.ResourceName, remapped.PlacedX!.Value));
        Assert.False(first.SourceImage.HasAuthoredRenderingIntent);
        Assert.True(second.SourceImage.HasAuthoredRenderingIntent);
        Assert.Equal(OfficeIccRenderingIntent.Perceptual, second.SourceImage.RenderingIntent);

        byte[] firstBytes = first.SourceImage.Bytes;
        Assert.Equal(new byte[] { 0, 255, 0, 0, 126, 0, 255, 0, 64 }, DecodeStoredPngIdat(firstBytes));
        Assert.Equal(firstBytes, second.SourceImage.Bytes);
        Assert.Equal(new byte[] { 0, 0, 0, 255, 126, 255, 255, 0, 64 }, DecodeStoredPngIdat(remapped.SourceImage.Bytes));

        using var destination = new BufferMutatingImageDestination();
        first.SourceImage.CopyTo(destination);
        Assert.Equal(firstBytes, destination.ToArray());
        Assert.Equal(firstBytes, first.SourceImage.Bytes);
        Assert.Equal(firstBytes, second.SourceImage.Bytes);
    }

    private static byte[] BuildRepeatedIndexedSoftMaskPdf() {
        const string redGreen = "FF000000FF00";
        const string blueYellow = "0000FFFFFF00";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 3 /Kids [3 0 R 4 0 R 5 0 R] >>\nendobj",
            Page(3, 6, "ImFirst", redGreen),
            Page(4, 7, "ImSecond", redGreen),
            Page(5, 8, "ImRemapped", blueYellow),
            Content(6, "q 24 0 0 24 36 160 cm /ImFirst Do Q"),
            Content(7, "q 24 0 0 24 72 160 cm /Perceptual ri /ImSecond Do Q"),
            Content(8, "q 24 0 0 24 108 160 cm /ImRemapped Do Q"),
            "9 0 obj\n<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /ColorSpace /Palette /BitsPerComponent 1 /SMask 10 0 R /Length 1 >>",
            "stream\n@\nendstream\nendobj",
            BuildSoftMaskObject(10, new byte[] { 126, 64 }, 2, 1),
            "trailer\n<< /Root 1 0 R /Size 11 >>",
            "%%EOF\n"
        });
        return Encoding.ASCII.GetBytes(pdf);

        static string Page(int page, int content, string resource, string palette) =>
            FormattableString.Invariant(
                $"{page} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 220 220] /Contents {content} 0 R /Resources << /ColorSpace << /Palette [/Indexed /DeviceRGB 1 <{palette}>] >> /XObject << /{resource} 9 0 R >> >> >>\nendobj");

        static string Content(int number, string content) =>
            number.ToString(CultureInfo.InvariantCulture) + " 0 obj\n<< /Length " +
            Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + content + "\nendstream\nendobj";
    }

    private sealed class BufferMutatingImageDestination : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            if (count > 0) buffer[offset] ^= 0xFF;
        }
    }
}
