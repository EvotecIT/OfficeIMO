namespace OfficeIMO.Pdf;

/// <summary>Output and proof produced by an existing-document encryption mutation.</summary>
public sealed class PdfSecurityMutationResult {
    private readonly byte[] _pdf;

    internal PdfSecurityMutationResult(
        PdfSecurityMutationKind kind,
        byte[] pdf,
        PdfMutationPlan mutationPlan,
        PdfRewritePreservationReport preservationReport,
        PdfDocumentSecurityInfo sourceSecurity,
        PdfDocumentSecurityInfo outputSecurity,
        PdfReadOptions? outputReadOptions) {
        Kind = kind;
        _pdf = (byte[])pdf.Clone();
        MutationPlan = mutationPlan;
        PreservationReport = preservationReport;
        SourceSecurity = sourceSecurity;
        OutputSecurity = outputSecurity;
        OutputReadOptions = outputReadOptions;
    }

    /// <summary>Security mutation that produced this result.</summary>
    public PdfSecurityMutationKind Kind { get; }

    /// <summary>Rewritten PDF bytes.</summary>
    public byte[] Pdf => (byte[])_pdf.Clone();

    /// <summary>Full-rewrite decision and required proof.</summary>
    public PdfMutationPlan MutationPlan { get; }

    /// <summary>Comparison proving preservation of supported non-security document structures.</summary>
    public PdfRewritePreservationReport PreservationReport { get; }

    /// <summary>Security state read from the source document.</summary>
    public PdfDocumentSecurityInfo SourceSecurity { get; }

    /// <summary>Security state read back from the rewritten document.</summary>
    public PdfDocumentSecurityInfo OutputSecurity { get; }

    /// <summary>True when the output is protected by Standard password security.</summary>
    public bool IsEncrypted => OutputSecurity.HasEncryption;

    internal PdfReadOptions? OutputReadOptions { get; }

    /// <summary>Opens the rewritten bytes through the normal fluent document API.</summary>
    public PdfDocument ToDocument() => PdfDocument.Open(_pdf, OutputReadOptions);
}
