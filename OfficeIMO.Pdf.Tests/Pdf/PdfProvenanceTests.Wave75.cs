using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void FilterlessRequiredDecodeReusesThePermissiveBudgetEntry() {
        var budget = new PdfDecodedStreamBudget(new PdfReadLimits {
            MaxDecodedStreamBytes = 8,
            MaxTotalDecodedStreamBytes = 5
        });
        var stream = new PdfStream(new PdfDictionary(), new byte[] { 1, 2, 3 });
        var objects = new Dictionary<int, PdfIndirectObject>();

        byte[] permissive = budget.Decode(stream, objects, maximumRequestedBytes: 8);
        byte[] required = budget.DecodeRequired(stream, objects, maximumRequestedBytes: 8);

        Assert.Same(permissive, required);
        Assert.Equal(3, budget.UsedBytes);
    }

    [Fact]
    public void XrefStreamTrailerPreservesBinaryIdentifierBytes() {
        byte[] expected = { 0x00, 0x80, 0xFF, 0xFE };
        byte[] pdf = BuildWave75XrefStreamPdf("0080FFFE");

        string trailer = PdfSyntax.ParseObjects(pdf).TrailerRaw;

        Assert.Equal(expected, PdfSyntax.ReadPermanentTrailerIdentifier(trailer));
        Assert.Contains("<0080FFFE>", trailer, StringComparison.Ordinal);
    }

    private static byte[] BuildWave75XrefStreamPdf(string identifierHex) {
        using var output = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }
        void WriteObject(int number, string body) {
            offsets[number] = checked((int)output.Position);
            WriteAscii(number + " 0 obj\n" + body + "\nendobj\n");
        }

        WriteAscii("%PDF-1.5\n");
        WriteObject(1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(2, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>");
        WriteObject(3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>");
        const int xrefObjectNumber = 4;
        offsets[xrefObjectNumber] = checked((int)output.Position);
        byte[] entries = new byte[5 * 7];
        WriteWave75XrefEntry(entries, 0, 0, 0, 65535);
        for (int number = 1; number <= xrefObjectNumber; number++) {
            WriteWave75XrefEntry(entries, number, 1, offsets[number], 0);
        }
        WriteAscii("4 0 obj\n<< /Type /XRef /Size 5 /Root 1 0 R /W [1 4 2] /Index [0 5] /ID [<" +
            identifierHex + "> <0102>] /Length " + entries.Length + " >>\nstream\n");
        output.Write(entries, 0, entries.Length);
        WriteAscii("\nendstream\nendobj\nstartxref\n" + offsets[xrefObjectNumber] + "\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteWave75XrefEntry(byte[] entries, int number, byte type, int field1, int field2) {
        int offset = number * 7;
        entries[offset] = type;
        entries[offset + 1] = (byte)(field1 >> 24);
        entries[offset + 2] = (byte)(field1 >> 16);
        entries[offset + 3] = (byte)(field1 >> 8);
        entries[offset + 4] = (byte)field1;
        entries[offset + 5] = (byte)(field2 >> 8);
        entries[offset + 6] = (byte)field2;
    }
}
