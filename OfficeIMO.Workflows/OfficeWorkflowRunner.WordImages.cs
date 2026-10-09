using System.Globalization;
using OfficeIMO.Drawing;
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
        foreach (WordImageOptimizationItem item in report.Images) AddWordImageDiagnostic(item, report.Applied, diagnostics);
        diagnostics.Add(new OfficeWorkflowDiagnostic("WordImageInventory",
            report.ImageCount + " unique embedded image(s); " + report.ExternalReferenceCount + " external reference(s) preserved.",
            stage: "images", details: new Dictionary<string, string> {
                ["requiredStagedBytes"] = report.RequiredStagedBytes.ToString(CultureInfo.InvariantCulture),
                ["applied"] = report.Applied.ToString()
            }));
        string summary = report.OptimizedCount + " image candidate(s); " + report.BytesSaved + " encoded media bytes saved" + (analyze ? " by analysis." : ".");
        if (analyze) return new OperationArtifact(null, summary, null);
        byte[] bytes;
        OfficeWorkflowConversionEvidence? evidence = null;
        string outputExtension = Path.GetExtension(request.OutputStream?.Name ?? request.OutputPath!);
        if (string.Equals(outputExtension, ".pdf", StringComparison.OrdinalIgnoreCase)) {
            // Media has already been optimized once. PDF generation keeps its default image policy.
            var conversion = word.ToPdfDocumentResult(new WordToPdfOptions(), token);
            (bytes, evidence) = SerializePdfConversion(conversion, request.Limits.MaximumOutputBytes, token);
            AddConversionDiagnostics(evidence, diagnostics);
        } else {
            using var output = new OfficeWorkflowBoundedMemoryStream(request.Limits.MaximumOutputBytes);
            WordFileFormat format = string.Equals(outputExtension, ".doc", StringComparison.OrdinalIgnoreCase)
                ? WordFileFormat.Doc : WordFileFormat.Docx;
            word.SaveAsync(output, format, new WordSaveOptions { SignedDocumentPolicy = options.SignedDocumentPolicy }, token).GetAwaiter().GetResult();
            bytes = output.ToArray();
        }
        return new OperationArtifact(bytes, summary, null, ConversionEvidence: evidence);
    }

    private static void AddWordImageDiagnostic(WordImageOptimizationItem item, bool applied,
        List<OfficeWorkflowDiagnostic> diagnostics) {
        bool removesMetadata = item.Metadata?.HasLoss == true ||
            (item.Metadata != null && item.Metadata.Stripped != OfficeImageMetadataKinds.None);
        var details = new Dictionary<string, string> {
            ["partUri"] = item.PartUri, ["references"] = item.ReferenceCount.ToString(CultureInfo.InvariantCulture),
            ["originalBytes"] = item.OriginalBytes.ToString(CultureInfo.InvariantCulture),
            ["finalBytes"] = item.FinalBytes.ToString(CultureInfo.InvariantCulture),
            ["originalFormat"] = item.Original.Format.ToString(), ["finalFormat"] = item.Final.Format.ToString(),
            ["originalWidth"] = item.Original.Width.ToString(CultureInfo.InvariantCulture),
            ["originalHeight"] = item.Original.Height.ToString(CultureInfo.InvariantCulture),
            ["finalWidth"] = item.Final.Width.ToString(CultureInfo.InvariantCulture),
            ["finalHeight"] = item.Final.Height.ToString(CultureInfo.InvariantCulture),
            ["status"] = item.Status.ToString(),
            ["applied"] = (applied && item.Status == WordImageOptimizationStatus.Optimized).ToString()
        };
        string metadata = "";
        if (item.Metadata != null) {
            details["candidateMetadataPolicy"] = item.Metadata.Policy.ToString();
            details["candidateMetadataSource"] = item.Metadata.Source.ToString();
            details["candidateMetadataPreserved"] = item.Metadata.Preserved.ToString();
            details["candidateMetadataLost"] = item.Metadata.Lost.ToString();
            details["candidateMetadataStripped"] = item.Metadata.Stripped.ToString();
            if (removesMetadata) metadata = " Candidate metadata: lost " + item.Metadata.Lost + "; stripped " + item.Metadata.Stripped + ".";
        }
        diagnostics.Add(new OfficeWorkflowDiagnostic("WordImage" + item.Status,
            item.PartUri + ": " + item.Status + "; " + item.BytesSaved.ToString(CultureInfo.InvariantCulture) + " encoded bytes saved." + metadata,
            removesMetadata ? OfficeWorkflowDiagnosticSeverity.Warning : OfficeWorkflowDiagnosticSeverity.Information,
            stage: "images", details: details));
    }
}
