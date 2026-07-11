using OfficeIMO.Rtf.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRawTextApiConsistencyTests {
    [Fact]
    public void RawRtf_UsesSourceExplicitPdfNames() {
        Assert.DoesNotContain(typeof(RtfPdfConverterExtensions).GetMethods(), method =>
            method.Name == "ToPdf" && method.GetParameters()[0].ParameterType == typeof(string));
        Assert.Contains(typeof(RtfPdfConverterExtensions).GetMethods(), method =>
            method.Name == "ToPdfFromRtf" && method.GetParameters()[0].ParameterType == typeof(string));
    }
}
