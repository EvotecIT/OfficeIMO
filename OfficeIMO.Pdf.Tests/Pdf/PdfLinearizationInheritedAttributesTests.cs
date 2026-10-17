using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfLinearizationInheritedAttributesTests {
    [Fact]
    public void WebProfile_PreservesEffectivePageValuesAndRemovesPageTreeInheritance() {
        const string content = "BT /F1 12 Tf 20 30 Td (Inherited resource) Tj ET";
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 4 /Kids [3 0 R 6 0 R] /Resources 7 0 R /MediaBox [0 0 200 300] /CropBox [1 2 190 290] /Rotate 90 >>", "endobj",
            "3 0 obj", "<< /Type /Pages /Parent 2 0 R /Count 2 /Kids [4 0 R 5 0 R] /Resources 8 0 R /MediaBox [0 0 400 500] /Rotate 180 >>", "endobj",
            "4 0 obj", "<< /Type /Page /Parent 3 0 R /Contents 12 0 R >>", "endobj",
            "5 0 obj", "<< /Type /Page /Parent 3 0 R /Resources 9 0 R /CropBox [10 20 300 400] /Rotate 270 /Contents 12 0 R >>", "endobj",
            "6 0 obj", "<< /Type /Pages /Parent 2 0 R /Count 2 /Kids [10 0 R 11 0 R] /CropBox [5 6 180 280] >>", "endobj",
            "7 0 obj", "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>", "endobj",
            "8 0 obj", "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Courier >> >> >>", "endobj",
            "9 0 obj", "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Times-Roman >> >> >>", "endobj",
            "10 0 obj", "<< /Type /Page /Parent 6 0 R /MediaBox [0 0 250 350] /Rotate 0 /Contents 12 0 R >>", "endobj",
            "11 0 obj", "<< /Type /Page /Parent 6 0 R /Contents 12 0 R >>", "endobj",
            "12 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 13 >>", "%%EOF"
        }));
        PdfReadDocument before = PdfReadDocument.Open(source);

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source, PdfOptimizationProfile.Web);

        Assert.True(result.CandidateLinearized);
        PdfReadDocument after = PdfReadDocument.Open(result.Bytes);
        var (outputObjects, _) = PdfSyntax.ParseObjects(result.Bytes);
        string[] inheritedKeys = { "Resources", "MediaBox", "CropBox", "Rotate" };
        string[] expectedFonts = { "Courier", "Times-Roman", "Helvetica", "Helvetica" };
        Assert.Equal(4, before.Pages.Count);
        Assert.Equal(before.Pages.Count, after.Pages.Count);
        for (int index = 0; index < before.Pages.Count; index++) {
            PdfReadPage original = before.Pages[index];
            PdfReadPage optimized = after.Pages[index];
            AssertBoxEqual(original.GetGeometry().MediaBox, optimized.GetGeometry().MediaBox);
            AssertBoxEqual(original.GetGeometry().CropBox, optimized.GetGeometry().CropBox);
            Assert.Equal(original.GetRotationDegrees(), optimized.GetRotationDegrees());
            Assert.Equal(original.GetPageSize(), optimized.GetPageSize());
            PdfTextSpan originalSpan = Assert.Single(original.GetTextSpans());
            PdfTextSpan optimizedSpan = Assert.Single(optimized.GetTextSpans());
            Assert.Equal(expectedFonts[index], originalSpan.BaseFont);
            Assert.Equal(originalSpan.BaseFont, optimizedSpan.BaseFont);
            Assert.Equal("Inherited resource", original.ExtractText().Trim());
            Assert.Equal(original.ExtractText(), optimized.ExtractText());

            PdfDictionary leaf = Assert.IsType<PdfDictionary>(outputObjects[optimized.ObjectNumber].Value);
            foreach (string key in inheritedKeys) Assert.True(leaf.Items.ContainsKey(key));
        }

        PdfDictionary[] pageTreeNodes = outputObjects.Values.Select(item => item.Value)
            .OfType<PdfDictionary>().Where(dictionary => dictionary.Get<PdfName>("Type")?.Name == "Pages").ToArray();
        Assert.NotEmpty(pageTreeNodes);
        foreach (PdfDictionary pageTree in pageTreeNodes) {
            foreach (string key in inheritedKeys) Assert.False(pageTree.Items.ContainsKey(key));
        }
    }

    private static void AssertBoxEqual(PdfPageBox? expected, PdfPageBox? actual) {
        Assert.NotNull(expected);
        Assert.NotNull(actual);
        Assert.Equal((expected.Left, expected.Bottom, expected.Right, expected.Top),
            (actual.Left, actual.Bottom, actual.Right, actual.Top));
    }
}
