using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
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
