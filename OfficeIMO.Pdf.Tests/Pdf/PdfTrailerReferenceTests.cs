using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTrailerReferenceTests {
    [Theory]
    [InlineData("trailer\n<< /Root 1 0 R /Info 2 0 R /Encrypt 3 0 R >>", 3, 1, 2)]
    [InlineData("trailer\n<< /Root null /Info 5 0 R >>\ntrailer\n<< /Root 1 0 R /Encrypt 7 0 R >>", 7, null, 5)]
    [InlineData("trailer\n<< /Encrypt /None >>\ntrailer\n<< /Encrypt 7 0 R /Root 1 0 R >>", null, 1, null)]
    public void CombinedTrailerScanMatchesSingleReferenceLookup(
        string trailerRaw,
        int? expectedEncrypt,
        int? expectedRoot,
        int? expectedInfo) {
        var references = PdfSyntax.ReadTrailerReferences(trailerRaw, "Encrypt", "Root", "Info", null);

        Assert.Equal(expectedEncrypt, references.First?.ObjectNumber);
        Assert.Equal(expectedRoot, references.Second?.ObjectNumber);
        Assert.Equal(expectedInfo, references.Third?.ObjectNumber);
        Assert.Equal(PdfSyntax.ReadTrailerReference(trailerRaw, "Encrypt")?.ObjectNumber, references.First?.ObjectNumber);
        Assert.Equal(PdfSyntax.ReadTrailerReference(trailerRaw, "Root")?.ObjectNumber, references.Second?.ObjectNumber);
        Assert.Equal(PdfSyntax.ReadTrailerReference(trailerRaw, "Info")?.ObjectNumber, references.Third?.ObjectNumber);
    }
}
