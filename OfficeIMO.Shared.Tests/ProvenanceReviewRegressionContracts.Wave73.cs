using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Theory]
    [InlineData("utf16-le")]
    [InlineData("utf16-be")]
    [InlineData("utf32-le")]
    [InlineData("utf32-be")]
    public void UnknownExtensionHtmlDetectionHonorsUnicodePreambles(string encodingName) {
        Encoding encoding = encodingName switch {
            "utf16-le" => new UnicodeEncoding(bigEndian: false, byteOrderMark: true),
            "utf16-be" => new UnicodeEncoding(bigEndian: true, byteOrderMark: true),
            "utf32-le" => new UTF32Encoding(bigEndian: false, byteOrderMark: true),
            "utf32-be" => new UTF32Encoding(bigEndian: true, byteOrderMark: true),
            _ => throw new InvalidOperationException("Unknown test encoding.")
        };
        const string html = "<!-- leading --><!doctype html><html><body>content</body></html>";
        byte[] bytes = encoding.GetPreamble().Concat(encoding.GetBytes(html)).ToArray();

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(bytes, "asset.bin");

        Assert.Equal(OfficeProvenanceAssetFormat.Html, report.Format);
    }

    [Theory]
    [InlineData("utf16-le")]
    [InlineData("utf16-be")]
    [InlineData("utf32-le")]
    [InlineData("utf32-be")]
    public void HtmlAssessmentHonorsUnicodePreamblesForTextIntegrity(string encodingName) {
        Encoding encoding = encodingName switch {
            "utf16-le" => new UnicodeEncoding(bigEndian: false, byteOrderMark: true),
            "utf16-be" => new UnicodeEncoding(bigEndian: true, byteOrderMark: true),
            "utf32-le" => new UTF32Encoding(bigEndian: false, byteOrderMark: true),
            "utf32-be" => new UTF32Encoding(bigEndian: true, byteOrderMark: true),
            _ => throw new InvalidOperationException("Unknown test encoding.")
        };
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".html");
        const string html = "<!doctype html><html><body>review\u200Bthis</body></html>";
        File.WriteAllBytes(path, encoding.GetPreamble().Concat(encoding.GetBytes(html)).ToArray());
        try {
            OfficeProvenanceAssessmentReport report = OfficeProvenanceAssessment.InspectFile(path);

            OfficeTextIntegrityFinding finding = Assert.Single(report.TextIntegrity!.Findings);
            Assert.Equal(OfficeTextIntegrityFindingKind.ZeroWidthSpace, finding.Kind);
            Assert.Equal(33, finding.TextOffset);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExtensionlessUtf32BigEndianSvgIsDetected(bool emitByteOrderMark) {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\"><metadata /></svg>";
        byte[] bytes = new UTF32Encoding(bigEndian: true, byteOrderMark: emitByteOrderMark)
            .GetPreamble()
            .Concat(new UTF32Encoding(bigEndian: true, byteOrderMark: false).GetBytes(svg))
            .ToArray();

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(bytes);

        Assert.Equal(OfficeProvenanceAssetFormat.Svg, report.Format);
    }
}
