using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfIdentityMemoryContractsTests {
    [Fact]
    public void RedactionImageGraphIdentityDoesNotMaterializeLargeStringPayloads() {
        var value = new PdfDictionary();
        value.Items["Payload"] = new PdfStringObj(new byte[4 * 1024 * 1024]);
        var identity = new StringBuilder();

        PdfRedactionImageIdentity.AppendObjectGraph(
            identity,
            value,
            new Dictionary<int, PdfIndirectObject>());

        Assert.InRange(identity.Length, 40, 80);
    }

    [Fact]
    public void DirectStreamIdentityIncludesLargeStringPayloadWithoutExpandingIt() {
        var firstDictionary = new PdfDictionary();
        firstDictionary.Items["Payload"] = new PdfStringObj(new byte[4 * 1024 * 1024]);
        var secondDictionary = new PdfDictionary();
        byte[] changed = new byte[4 * 1024 * 1024];
        changed[changed.Length - 1] = 1;
        secondDictionary.Items["Payload"] = new PdfStringObj(changed);

        int first = PdfDirectStreamIdentity.Compute(new PdfStream(firstDictionary, new byte[] { 1 }));
        int second = PdfDirectStreamIdentity.Compute(new PdfStream(secondDictionary, new byte[] { 1 }));

        Assert.NotEqual(first, second);
    }
}
