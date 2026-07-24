using System.Reflection;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAllSeverityBatch17SecurityTests {
    [Fact]
    public void SignaturePlaceholderScanIgnoresDecoysInsideStreamBodies() {
        string streamPayload =
            "BT\n" +
            "endstream\n<< /Type /Sig /ByteRange [0 0 0 0] /Contents <0000> >>\nendobj\n" +
            "ET\n";
        byte[] source = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Length " + Encoding.ASCII.GetByteCount(streamPayload) + " >>\nstream\n" +
            streamPayload +
            "endstream\nendobj\n" +
            "2 0 obj\n<< /Type /Sig /ByteRange [0 0 0 0] /Contents <0000> >>\nendobj\n");
        MethodInfo method = typeof(PdfIncrementalUpdater).GetMethod(
            "FindZeroFilledSignatureContents",
            BindingFlags.Static | BindingFlags.NonPublic)!;
        object[] arguments = { source, 0, 0, 0 };

        int count = (int)method.Invoke(null, arguments)!;

        Assert.Equal(1, count);
        Assert.Equal(2, (int)arguments[3]);
        Assert.True((int)arguments[1] > 0);
        Assert.Equal(4, (int)arguments[2]);
    }

    [Fact]
    public void RawSignatureCompletionIgnoresEndstreamDecoyInDeclaredStreamData() {
        const string decoy =
            "endstream\n<< /Type /Sig /ByteRange [0 0 0 0] /Contents <0000> >>\nendobj";
        byte[] source = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Paragraph(paragraph => paragraph.Text(decoy))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions { ReservedSignatureContentsBytes = 256 });

        byte[] completed = PdfIncrementalUpdater.ApplyExternalSignature(
            preparation.PreparedPdf,
            new byte[] { 0x30, 0x01, 0x00 });

        Assert.Equal(preparation.PreparedPdf.Length, completed.Length);
        Assert.Contains("300100", Encoding.ASCII.GetString(completed), StringComparison.Ordinal);
    }

    [Fact]
    public void SignaturePlaceholderScanResolvesIndirectStreamLengthBeforeStructuralFallback() {
        string streamPayload =
            "BT\n" +
            "endstream\nendobj\n" +
            "9 0 obj\n<< /Type /Sig /ByteRange [0 0 0 0] /Contents <0000> >>\nendobj\n" +
            "ET\n";
        byte[] source = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Length 5 0 R >>\nstream\n" +
            streamPayload +
            "endstream\nendobj\n" +
            "5 0 obj\n" + Encoding.ASCII.GetByteCount(streamPayload) + "\nendobj\n" +
            "2 0 obj\n<< /Type /Sig /ByteRange [0 0 0 0] /Contents <0000> >>\nendobj\n");
        MethodInfo method = typeof(PdfIncrementalUpdater).GetMethod(
            "FindZeroFilledSignatureContents",
            BindingFlags.Static | BindingFlags.NonPublic)!;
        object[] arguments = { source, 0, 0, 0 };

        int count = (int)method.Invoke(null, arguments)!;

        Assert.Equal(1, count);
        Assert.Equal(2, (int)arguments[3]);
    }
}
