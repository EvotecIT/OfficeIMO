using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

internal sealed class PdfInteroperabilityCorpusCase {
    public PdfInteroperabilityCorpusCase(
        string id,
        byte[] pdf,
        PdfMutationExecutionMode expectedMetadataMode,
        PdfReadOptions? readOptions,
        params string[] features) {
        Id = id;
        Pdf = pdf;
        ExpectedMetadataMode = expectedMetadataMode;
        ReadOptions = readOptions;
        Features = Array.AsReadOnly(features);
    }

    public string Id { get; }

    public byte[] Pdf { get; }

    public PdfMutationExecutionMode ExpectedMetadataMode { get; }

    public PdfReadOptions? ReadOptions { get; }

    public IReadOnlyList<string> Features { get; }
}

internal static class PdfInteroperabilityCorpusSupport {
    public const int CaseCount = 12;

    public static IReadOnlyList<PdfInteroperabilityCorpusCase> Build() {
        byte[] encrypted = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Corpus encrypted"))
            .ToBytes();

        return new[] {
            new PdfInteroperabilityCorpusCase(
                "classic-xref",
                PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Corpus classic xref")).ToBytes(),
                PdfMutationExecutionMode.FullRewrite,
                null,
                "classic-xref", "born-digital"),
            new PdfInteroperabilityCorpusCase(
                "xref-and-object-streams",
                PdfRewritePreservationTestSupport.BuildSourceStructurePreservationProofPdf(),
                PdfMutationExecutionMode.AppendOnly,
                null,
                "xref-stream", "object-stream", "incremental"),
            new PdfInteroperabilityCorpusCase(
                "hybrid-reference",
                PdfExternalDocumentCompatibilityTests.BuildHybridClassicXrefPdfWithXRefStmAndTrailingStaleDuplicatePage(),
                PdfMutationExecutionMode.AppendOnly,
                null,
                "classic-xref", "xref-stream", "hybrid-reference"),
            new PdfInteroperabilityCorpusCase(
                "unusual-generations-incremental",
                PdfExternalDocumentCompatibilityTests.BuildIncrementalClassicXrefPdfWithWrongGenerationReplacementPage(),
                PdfMutationExecutionMode.AppendOnly,
                null,
                "classic-xref", "unusual-generation", "incremental"),
            new PdfInteroperabilityCorpusCase(
                "linearized-marker",
                BuildLinearizedPdf(),
                PdfMutationExecutionMode.FullRewrite,
                null,
                "classic-xref", "linearized"),
            new PdfInteroperabilityCorpusCase(
                "signed-certified",
                PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf(),
                PdfMutationExecutionMode.Blocked,
                null,
                "signature", "doc-mdp", "field-mdp", "dss", "incremental"),
            new PdfInteroperabilityCorpusCase(
                "password-encrypted",
                encrypted,
                PdfMutationExecutionMode.FullRewrite,
                new PdfReadOptions { Password = "open" },
                "encrypted", "standard-security", "password"),
            new PdfInteroperabilityCorpusCase(
                "tagged-structure",
                PdfRewritePreservationTestSupport.BuildTaggedPreservationProofPdf(),
                PdfMutationExecutionMode.AppendOnly,
                null,
                "tagged-content", "structure-tree"),
            new PdfInteroperabilityCorpusCase(
                "optional-content",
                PdfOptionalContentSupport.BuildOptionalContentMetadataPdf(),
                PdfMutationExecutionMode.FullRewrite,
                null,
                "optional-content", "layers"),
            new PdfInteroperabilityCorpusCase(
                "catalog-rich",
                PdfRewritePreservationTestSupport.BuildPreservationProofPdf(),
                PdfMutationExecutionMode.FullRewrite,
                null,
                "attachments", "name-trees", "output-intents", "xmp", "outlines"),
            new PdfInteroperabilityCorpusCase(
                "active-content",
                PdfRewritePreservationMatrixScenarioSupport.BuildPageActiveContentProofPdf(),
                PdfMutationExecutionMode.Blocked,
                null,
                "active-content", "javascript", "page-actions"),
            new PdfInteroperabilityCorpusCase(
                "complex-acroform",
                PdfFormAppearanceProofTestSupport.BuildFormAppearanceProofPdf(),
                PdfMutationExecutionMode.Blocked,
                null,
                "acroform", "widget-annotations", "appearances", "choice-field")
        };
    }

    private static byte[] BuildLinearizedPdf() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, long>();
        Write(stream, "%PDF-1.7\n");
        WriteObject(stream, offsets, 1, "<< /Linearized 1 /L 0000000000 /H [0 0] /O 4 /E 0 /N 1 /T 0 >>");
        WriteObject(stream, offsets, 2, "<< /Type /Catalog /Pages 3 0 R >>");
        WriteObject(stream, offsets, 3, "<< /Type /Pages /Count 1 /Kids [4 0 R] >>");
        WriteObject(stream, offsets, 4, "<< /Type /Page /Parent 3 0 R /MediaBox [0 0 200 200] /Contents 5 0 R >>");
        WriteObject(stream, offsets, 5, "<< /Length 0 >>\nstream\n\nendstream");
        long xref = stream.Position;
        Write(stream, "xref\n0 6\n0000000000 65535 f \n");
        for (int i = 1; i <= 5; i++) {
            Write(stream, offsets[i].ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
        }

        Write(stream, "trailer\n<< /Size 6 /Root 2 0 R >>\nstartxref\n" + xref.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        byte[] pdf = stream.ToArray();
        string length = pdf.Length.ToString("D10", CultureInfo.InvariantCulture);
        byte[] marker = Encoding.ASCII.GetBytes("0000000000");
        int index = IndexOf(pdf, marker);
        Encoding.ASCII.GetBytes(length).CopyTo(pdf, index);
        return pdf;
    }

    private static void WriteObject(Stream stream, Dictionary<int, long> offsets, int objectNumber, string body) {
        offsets.Add(objectNumber, stream.Position);
        Write(stream, objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + body + "\nendobj\n");
    }

    private static void Write(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static int IndexOf(byte[] source, byte[] value) {
        for (int i = 0; i <= source.Length - value.Length; i++) {
            if (source.AsSpan(i, value.Length).SequenceEqual(value)) {
                return i;
            }
        }

        throw new InvalidOperationException("Expected corpus marker was not found.");
    }
}
