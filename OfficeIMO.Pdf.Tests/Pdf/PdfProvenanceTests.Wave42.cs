using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ProvenanceLimitClonePreservesCallerSpecificLimits() {
        var source = new PdfReadLimits {
            MaxRetainedContentBytes = 123,
            MaxTextSearchMatches = 17,
            MaxWidgetActions = 7
        };

        PdfReadLimits effective = source.WithMaximumContainerEntries(10);

        Assert.Equal(123, effective.MaxRetainedContentBytes);
        Assert.Equal(17, effective.MaxTextSearchMatches);
        Assert.Equal(7, effective.MaxWidgetActions);
    }

    [Fact]
    public void CatalogJavaScriptStreamsShareTheProvenanceExpandedByteBudget() {
        byte[] pdf = BuildDecodedStreamBudgetPdf(catalogScripts: 2, includeXmp: false, includeWidget: false);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfProvenance.Inspect(
            pdf,
            new OfficeProvenanceOptions { MaxExpandedContainerBytes = 100 }));

        Assert.Equal(PdfReadLimitKind.TotalDecodedStreamBytes, exception.Kind);
    }

    [Fact]
    public void CatalogXmpSharesTheProvenanceExpandedByteBudget() {
        byte[] pdf = BuildDecodedStreamBudgetPdf(catalogScripts: 1, includeXmp: true, includeWidget: false);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfProvenance.Inspect(
            pdf,
            new OfficeProvenanceOptions { MaxExpandedContainerBytes = 100 }));

        Assert.Equal(PdfReadLimitKind.TotalDecodedStreamBytes, exception.Kind);
    }

    [Fact]
    public void WidgetJavaScriptSharesTheProvenanceExpandedByteBudget() {
        byte[] pdf = BuildDecodedStreamBudgetPdf(catalogScripts: 0, includeXmp: true, includeWidget: true);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfProvenance.Inspect(
            pdf,
            new OfficeProvenanceOptions { MaxExpandedContainerBytes = 100 }));

        Assert.Equal(PdfReadLimitKind.TotalDecodedStreamBytes, exception.Kind);
    }

    [Fact]
    public void ActiveAcroFormXfaGraphCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((_, catalog, candidate) => {
            var acroForm = new PdfDictionary();
            acroForm.Items["XFA"] = candidate;
            catalog.Items["AcroForm"] = acroForm;
        });
    }

    [Fact]
    public void MalformedUnicodeEmbeddedFileVariantDoesNotHideTheCandidate() {
        byte[] original = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] malformed = PdfDocumentObjectGraphRewriter.Rewrite(original, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, candidate));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpecification.Items["EF"]));
            embeddedFiles.Items["UF"] = PdfNull.Instance;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(malformed);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    private static byte[] BuildDecodedStreamBudgetPdf(int catalogScripts, bool includeXmp, bool includeWidget) {
        byte[] payload = Encoding.ASCII.GetBytes(new string('x', 60));
        byte[] compressed = CompressForWave42(payload);
        using var output = new MemoryStream();
        void Write(string value) {
            byte[] bytes = Encoding.ASCII.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }
        void WriteStream(int objectNumber) {
            Write(objectNumber + " 0 obj\n<< /Length " + compressed.Length + " /Filter /FlateDecode >>\nstream\n");
            output.Write(compressed, 0, compressed.Length);
            Write("\nendstream\nendobj\n");
        }

        var catalog = new StringBuilder("1 0 obj\n<< /Type /Catalog /Pages 2 0 R");
        if (includeWidget) catalog.Append(" /AcroForm 5 0 R");
        if (includeXmp) catalog.Append(" /Metadata 8 0 R");
        if (catalogScripts > 0) {
            catalog.Append(" /Names << /JavaScript << /Names [");
            for (int index = 0; index < catalogScripts; index++) {
                catalog.Append("(script").Append(index).Append(") ").Append(9 + (index * 2)).Append(" 0 R ");
            }
            catalog.Append("] >> >>");
        }
        catalog.Append(" >>\nendobj\n");
        Write("%PDF-1.7\n");
        Write(catalog.ToString());
        Write("2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        Write("3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100]");
        if (includeWidget) Write(" /Annots [6 0 R]");
        Write(" >>\nendobj\n");
        if (includeWidget) {
            Write("5 0 obj\n<< /Fields [6 0 R] >>\nendobj\n");
            Write("6 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Btn /T (run) /Rect [0 0 10 10] /A << /S /JavaScript /JS 7 0 R >> >>\nendobj\n");
            WriteStream(7);
        }
        if (includeXmp) WriteStream(8);
        for (int index = 0; index < catalogScripts; index++) {
            int actionNumber = 9 + (index * 2);
            Write(actionNumber + " 0 obj\n<< /S /JavaScript /JS " + (actionNumber + 1) + " 0 R >>\nendobj\n");
            WriteStream(actionNumber + 1);
        }
        Write("trailer\n<< /Root 1 0 R /Size 20 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] CompressForWave42(byte[] data) {
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }
        uint first = 1;
        uint second = 0;
        for (int index = 0; index < data.Length; index++) {
            first = (first + data[index]) % 65521;
            second = (second + first) % 65521;
        }
        uint checksum = (second << 16) | first;
        output.WriteByte((byte)(checksum >> 24));
        output.WriteByte((byte)(checksum >> 16));
        output.WriteByte((byte)(checksum >> 8));
        output.WriteByte((byte)checksum);
        return output.ToArray();
    }
}
