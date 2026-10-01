using System.Globalization;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OperationArtifact OptimizeWordImages(ValidatedRequest request,
        List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        byte[] input = ReadInput(request.InputPath, request.Limits, token);
        using var source = new MemoryStream(input, writable: false);
        WordLoadOptions load = CreateWordLoadOptions(request.Limits);
        bool analyze = request.Operation == OfficeWorkflowOperation.AnalyzeWordImages;
        load.AccessMode = analyze ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite;
        using WordDocument word = WordDocument.LoadAsync(source, load, cancellationToken: token).GetAwaiter().GetResult();
        if (word.SourceFormat == WordFileFormat.Doc &&
            (word.LegacyDocUnsupportedFeatures.Count != 0 || word.LegacyDocPreservedFeatures.Count != 0 || word.LegacyDocCompoundFeatures.Count != 0)) {
            if (!analyze) throw new NotSupportedException("Image optimization cannot publish a copy of legacy DOC content that was not fully projected. Inspect the legacy import report first.");
            diagnostics.Add(new OfficeWorkflowDiagnostic("LegacyDocProjectionIncomplete",
                "The image inventory includes projected legacy pictures only. Preserve-only legacy content remains in the original source.",
                OfficeWorkflowDiagnosticSeverity.Warning, "images"));
        }
        WordImageOptimizationOptions options = request.WordImageOptimization!;
        WordImageOptimizationReport report = analyze ? word.AnalyzeImageOptimization(options, token) : word.OptimizeImages(options, token);
        foreach (WordImageOptimizationItem item in report.Images) {
            diagnostics.Add(new OfficeWorkflowDiagnostic("WordImage" + item.Status,
                item.PartUri + ": " + item.Status + "; " + item.BytesSaved.ToString(CultureInfo.InvariantCulture) + " encoded bytes saved.",
                stage: "images", details: new Dictionary<string, string> {
                    ["partUri"] = item.PartUri, ["references"] = item.ReferenceCount.ToString(CultureInfo.InvariantCulture),
                    ["originalBytes"] = item.OriginalBytes.ToString(CultureInfo.InvariantCulture),
                    ["finalBytes"] = item.FinalBytes.ToString(CultureInfo.InvariantCulture),
                    ["status"] = item.Status.ToString()
                }));
        }
        diagnostics.Add(new OfficeWorkflowDiagnostic("WordImageInventory",
            report.ImageCount + " unique embedded image(s); " + report.ExternalReferenceCount + " external reference(s) preserved.", stage: "images"));
        string summary = report.OptimizedCount + " image candidate(s); " + report.BytesSaved + " encoded media bytes saved" + (analyze ? " by analysis." : ".");
        if (analyze) return new OperationArtifact(null, summary, null);
        byte[] bytes;
        string outputExtension = Path.GetExtension(request.OutputStream?.Name ?? request.OutputPath!);
        if (string.Equals(outputExtension, ".pdf", StringComparison.OrdinalIgnoreCase)) {
            // Media has already been optimized once. PDF generation keeps its default image policy.
            var conversion = word.ToPdfDocumentResult(new WordToPdfOptions(), token);
            bytes = SerializePdfConversion(conversion, request.Limits.MaximumOutputBytes, token);
            AddPdfWarnings(conversion.Warnings, diagnostics);
        } else {
            using var output = new OfficeWorkflowBoundedMemoryStream(request.Limits.MaximumOutputBytes);
            WordFileFormat format = string.Equals(outputExtension, ".doc", StringComparison.OrdinalIgnoreCase)
                ? WordFileFormat.Doc : WordFileFormat.Docx;
            word.SaveAsync(output, format, new WordSaveOptions { SignedDocumentPolicy = options.SignedDocumentPolicy }, token).GetAwaiter().GetResult();
            bytes = output.ToArray();
        }
        return new OperationArtifact(bytes, summary, null);
    }
}
