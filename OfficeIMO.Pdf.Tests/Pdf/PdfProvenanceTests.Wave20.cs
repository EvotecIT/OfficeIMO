using System.IO.Compression;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProvenanceStructuralBudgetCountsContainerMembers(bool useArray) {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] oversized = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            if (useArray) {
                var array = new PdfArray();
                for (int index = 0; index < 512; index++) array.Items.Add(PdfNull.Instance);
                catalog.Items["Wave20Budget"] = array;
            } else {
                var dictionary = new PdfDictionary();
                for (int index = 0; index < 512; index++) {
                    dictionary.Items["K" + index.ToString(System.Globalization.CultureInfo.InvariantCulture)] = PdfNull.Instance;
                }
                catalog.Items["Wave20Budget"] = dictionary;
            }
            return security.InfoObjectNumber;
        });
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 256 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => PdfProvenance.Inspect(oversized, options));

        Assert.Contains("container entry limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ActiveSubtypeLessAnnotationCannotMasqueradeAsTheFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] annotationCarrier = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            fileSpecification.Items.Remove("Subtype");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values.Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            var annotations = new PdfArray();
            annotations.Items.Add(candidate);
            page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(annotationCarrier);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(annotationCarrier);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(annotationCarrier, result.ToArray());
    }

    [Fact]
    public void ProvenanceExpandedContainerLimitAppliesWhileOpeningPdfStreams() {
        byte[] pdf = BuildWave20CompressedXrefStreamPdf();
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = pdf.LongLength + 1L,
            MaxManifestBytes = 16,
            MaxExpandedContainerBytes = 16
        };
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = 1024 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfProvenance.Inspect(pdf, options, readOptions));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(16, exception.Limit);
    }

    private static byte[] BuildWave20CompressedXrefStreamPdf() {
        byte[] decoded = new byte[70];
        byte[] encoded;
        using (var compressed = new MemoryStream()) {
            using (var compressor = new DeflateStream(compressed, CompressionLevel.Optimal, leaveOpen: true)) {
                compressor.Write(decoded, 0, decoded.Length);
            }
            encoded = compressed.ToArray();
        }

        using var output = new MemoryStream();
        void Write(string value) {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        Write("%PDF-1.5\n");
        Write("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        Write("2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        Write("3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>\nendobj\n");
        int xrefOffset = checked((int)output.Position);
        Write("5 0 obj\n<< /Type /XRef /Size 10 /Root 1 0 R /W [1 4 2] /Index [0 10] /Filter /FlateDecode /Length " + encoded.Length + " >>\nstream\n");
        output.Write(encoded, 0, encoded.Length);
        Write("\nendstream\nendobj\nstartxref\n" + xrefOffset + "\n%%EOF\n");
        return output.ToArray();
    }
}
