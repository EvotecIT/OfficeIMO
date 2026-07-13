using OfficeIMO.Rtf.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRawTextApiConsistencyTests {
    [Fact]
    public void RawRtfInputs_AreOwnedByTheNativeRtfLoader() {
        Assert.DoesNotContain(typeof(RtfPdfConverterExtensions).GetMethods(), method =>
            method.GetParameters().FirstOrDefault()?.ParameterType is Type sourceType
            && (sourceType == typeof(string) || sourceType == typeof(byte[]) || sourceType == typeof(Stream)));
    }
}
