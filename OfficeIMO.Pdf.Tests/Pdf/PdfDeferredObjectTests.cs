using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Objects held in object streams are validated when a document opens and built on first use.
/// These tests pin what callers observe at the reader boundary: one instance per object, and
/// load-time diagnostics and limits that do not depend on which objects are later read.
/// </summary>
public class PdfDeferredObjectTests {
    private const string Catalog = "<< /Type /Catalog /Pages 2 0 R >>";
    private const string Pages = "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>";
    private const string Page = "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>";
    private const string Font = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>";

    [Fact]
    public void CompressedObjectReadFromManyThreadsIsOneInstance() {
        byte[] pdf = BuildXrefStreamPdf((6, "<< /Kind /Unread /Values [1 2 3] /Next 5 0 R >>"));

        PdfReadDocument document = PdfReadDocument.Open(pdf);
        Assert.Equal("Deferred page", document.ExtractText().Trim());
        Assert.Empty(document.RepairReport.Diagnostics);

        // The object has not been read yet, so the threads below race to build it.
        Assert.False(document.Objects[6].IsValueBuilt);
        var values = new PdfObject[16];
        Parallel.For(0, values.Length, index => values[index] = document.Objects[6].Value);

        var dictionary = Assert.IsType<PdfDictionary>(values[0]);
        Assert.Equal("Unread", dictionary.Get<PdfName>("Kind")?.Name);
        Assert.All(values, value => Assert.Same(values[0], value));
        Assert.Same(values[0], document.Objects[6].Value);
    }

    [Fact]
    public void DamagedAndOrphanedCompressedObjectsAreReportedAtLoad() {
        byte[] pdf = BuildXrefStreamPdf(
            (6, "<< /Type /Annot /Subtype /Link /Rect [0 0 10 10] >>"),
            (7, "<< /Broken /Rect [0 0 10 10 >>"));

        PdfReadDocument document = PdfReadDocument.Open(pdf);

        Assert.Contains(document.RepairReport.Diagnostics, diagnostic =>
            diagnostic.Code == "unreadable-indirect-object" && diagnostic.ObjectNumber == 7);
        PdfRepairDiagnostic orphaned = Assert.Single(document.RepairReport.Diagnostics, diagnostic =>
            diagnostic.Code == "orphaned-semantic-objects");
        Assert.Equal(6, orphaned.ObjectNumber);

        PdfParseException strict = Assert.Throws<PdfParseException>(() =>
            PdfReadDocument.Open(pdf, new PdfLoadOptions { ParsingMode = PdfParsingMode.Strict }));
        Assert.Equal(7, strict.ObjectNumber);
    }

    [Fact]
    public void CompressedObjectLimitsAreEnforcedWhenTheDocumentOpens() {
        byte[] pdf = BuildXrefStreamPdf((6, "<< /Deep [[[[[[1]]]]]] >>"));
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxObjectNestingDepth = 3 } };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(pdf, options));

        Assert.Equal(PdfReadLimitKind.ObjectNestingDepth, exception.Kind);
    }

    private static byte[] BuildXrefStreamPdf(params (int ObjectNumber, string Body)[] extraObjects) {
        var packed = new List<(int ObjectNumber, string Body)> { (1, Catalog), (2, Pages), (3, Page), (5, Font) };
        packed.AddRange(extraObjects);

        var header = new StringBuilder();
        var body = new StringBuilder();
        foreach (var (objectNumber, objectBody) in packed) {
            header.Append(objectNumber.ToString(CultureInfo.InvariantCulture)).Append(' ')
                .Append(body.Length.ToString(CultureInfo.InvariantCulture)).Append(' ');
            body.Append(objectBody).Append('\n');
        }
        string objectStream = header.ToString() + body;

        using var output = new MemoryStream();
        var offsets = new Dictionary<int, int>();
        Write(output, "%PDF-1.5\n");
        string content = "BT /F1 12 Tf 72 720 Td (Deferred page) Tj ET";
        offsets[4] = (int)output.Position;
        Write(output, "4 0 obj\n<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream\nendobj\n");
        offsets[10] = (int)output.Position;
        Write(output, "10 0 obj\n<< /Type /ObjStm /N " + packed.Count.ToString(CultureInfo.InvariantCulture) +
            " /First " + header.Length.ToString(CultureInfo.InvariantCulture) +
            " /Length " + objectStream.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + objectStream + "\nendstream\nendobj\n");

        const int xrefObjectNumber = 11;
        const int size = xrefObjectNumber + 1;
        offsets[xrefObjectNumber] = (int)output.Position;
        var rows = new byte[size * 7];
        for (int index = 0; index < packed.Count; index++) {
            WriteRow(rows, packed[index].ObjectNumber, 2, 10, index);
        }
        foreach (int objectNumber in new[] { 4, 10, xrefObjectNumber }) {
            WriteRow(rows, objectNumber, 1, offsets[objectNumber], 0);
        }
        Write(output, xrefObjectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n<< /Type /XRef /Size " +
            size.ToString(CultureInfo.InvariantCulture) + " /Root 1 0 R /W [1 4 2] /Length " +
            rows.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(rows, 0, rows.Length);
        Write(output, "\nendstream\nendobj\nstartxref\n" + offsets[xrefObjectNumber].ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteRow(byte[] rows, int objectNumber, int type, int field1, int field2) {
        int row = objectNumber * 7;
        rows[row] = (byte)type;
        rows[row + 1] = (byte)(field1 >> 24);
        rows[row + 2] = (byte)(field1 >> 16);
        rows[row + 3] = (byte)(field1 >> 8);
        rows[row + 4] = (byte)field1;
        rows[row + 5] = (byte)(field2 >> 8);
        rows[row + 6] = (byte)field2;
    }

    private static void Write(Stream stream, string text) {
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }
}
