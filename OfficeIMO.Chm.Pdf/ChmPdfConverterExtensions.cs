using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Chm;

/// <summary>Renders each selected topic independently and merges its pages through the first-party PDF engine.</summary>
public static partial class ChmPdfConverterExtensions {
    // Topics render through the asynchronous archive resolver; the synchronous form waits on it off the caller's context.
    internal static PdfDocumentConversionResult ConvertToPdfWithEvidence(this ChmDocument document, ChmPdfExportOptions? options,
        CancellationToken cancellationToken) =>
        Task.Run(() => document.ConvertToPdfWithEvidenceAsync(options, cancellationToken), cancellationToken).GetAwaiter().GetResult();

    internal static async Task<PdfDocumentConversionResult> ConvertToPdfWithEvidenceAsync(this ChmDocument document, ChmPdfExportOptions? exportOptions,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(document);
        ChmConversionOptions? options = exportOptions?.ConversionOptions;
        HtmlPdfExportOptions? pdfOptions = exportOptions?.PdfOptions;
        ChmConversionOptions configured = options?.Clone() ?? new ChmConversionOptions();
        IReadOnlyList<ChmTopic> topics = document.SelectTopics(configured);
        var diagnostics = document.SourceDiagnostics();
        HtmlPdfExportOptions template = pdfOptions?.ClonePdf() ?? new HtmlPdfExportOptions();
        if (topics.Count > 1 && template.PdfOptions.TaggedStructureMode != PdfTaggedStructureMode.None) {
            // The PDF owner blocks full rewrites of tagged structure. Preserve topic CSS and page
            // boundaries through an explicitly reported untagged merge instead of bypassing that gate.
            template.PdfOptions.SetTaggedStructureMode(PdfTaggedStructureMode.None);
            diagnostics.Add(new global::OfficeIMO.OfficeDiagnostic("CHM_PDF_TAGGED_STRUCTURE_OMITTED", global::OfficeIMO.OfficeDiagnosticSeverity.Warning, "Multi-topic PDF uses untagged pages because the PDF owner cannot preserve a merged structure tree. Select one topic for tagged output.", null, "OfficeIMO.Chm.Pdf", OfficeConversionLossKind.Omission));
        }
        PdfStandardEncryptionOptions? encryption = template.PdfOptions.Encryption;
        template.PdfOptions.ClearEncryption();
        var pdfs = new List<PdfDocument>();
        long characters = 0, retainedBytes = 0;
        int nodes = 0;
        foreach (ChmTopic topic in topics) {
            cancellationToken.ThrowIfCancellationRequested();
            string html = topic.ReadHtml(cancellationToken); ChmDocument.ReserveCharacters(ref characters, html.Length, configured);
            HtmlPdfExportOptions rendering = template.ClonePdf();
            document.ConfigureRenderOptions(rendering, topic.Path);
            rendering.EmbeddedPackageResourceResolver = rendering.ResourcePolicy.AllowEmbeddedPackageResources ? document.CreateResourceResolver() : null;
            rendering.ResourceResolver = rendering.EmbeddedPackageResourceResolver;
            PdfDocumentConversionResult converted = await document.ParseConversionTopic(html, topic.Path, configured, ref nodes, cancellationToken)
                .ConvertToPdfWithEvidenceAsync(rendering, cancellationToken).ConfigureAwait(false);
            // Bound retained merge inputs before accumulating the entire publication.
            retainedBytes += converted.ToBytes(cancellationToken).LongLength;
            ChmDocument.EnforceOutput(retainedBytes, configured);
            pdfs.Add(converted.Value);
            diagnostics.AddRange(converted.Diagnostics.Select(item => new global::OfficeIMO.OfficeDiagnostic(item.Code, item.LossKind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning, item.Message, topic.Path, item.Format, item.LossKind)));
        }
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocument value = pdfs[0];
        if (pdfs.Count > 1) {
            PdfMergeResult merged = PdfDocument.MergeResult(new PdfMergeOptions {
                Policy = new PdfMergePolicy { Outlines = PdfMergeStructureMode.Combine, NamedDestinations = PdfMergeStructureMode.Combine }
            }, pdfs, cancellationToken);
            value = merged.RequireValue();
            foreach (PdfMergeDecision decision in merged.Report.Decisions.Where(item => item.DroppedCount != 0))
                diagnostics.Add(new global::OfficeIMO.OfficeDiagnostic("CHM_PDF_MERGE_OMISSION", global::OfficeIMO.OfficeDiagnosticSeverity.Warning, decision.Action, decision.Structure, "OfficeIMO.Pdf", OfficeConversionLossKind.Omission));
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (encryption != null) {
            PdfSecurityMutationResult encrypted = PdfSecurityEditor.Encrypt(new PdfDocumentConversionResult(value, new PdfConversionReport()).ToBytes(cancellationToken),
                encryption, maximumOutputBytes: configured.MaxOutputBytes, cancellationToken: cancellationToken);
            if (!encrypted.PreservationReport.IsPreserved) throw new InvalidOperationException("CHM PDF encryption did not preserve the merged document.");
            value = encrypted.ToDocument();
        }
        if (topics.Count > 1) diagnostics.Add(new global::OfficeIMO.OfficeDiagnostic("CHM_PDF_TOPIC_LINKS", global::OfficeIMO.OfficeDiagnosticSeverity.Warning, "Links between separate help topics are not rebuilt as PDF destinations.", null, "OfficeIMO.Chm.Pdf", OfficeConversionLossKind.Omission));
        if (document.Index.Count != 0) diagnostics.Add(new global::OfficeIMO.OfficeDiagnostic("CHM_PDF_INDEX", global::OfficeIMO.OfficeDiagnosticSeverity.Warning, "The compiled help keyword index is not recreated in PDF.", null, "OfficeIMO.Chm.Pdf", OfficeConversionLossKind.Omission));
        var result = new PdfDocumentConversionResult(value, new PdfConversionReport())
            .WithSourceConversionReport(new ChmConversionReport(topics.Select(topic => topic.Path), diagnostics));
        return result;
    }

}
