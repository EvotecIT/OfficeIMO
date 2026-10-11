using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Chm;

/// <summary>Renders each selected topic independently and merges its pages through the first-party PDF engine.</summary>
public static class ChmPdfConverterExtensions {
    /// <summary>Converts help topics to PDF, preserving per-topic CSS isolation and composing source and renderer reports.</summary>
    public static PdfDocumentConversionResult ToPdfDocumentResult(this ChmDocument document, ChmConversionOptions? options = null,
        HtmlToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) =>
        Task.Run(() => document.ToPdfDocumentResultAsync(options, pdfOptions, cancellationToken), cancellationToken).GetAwaiter().GetResult();

    /// <summary>Asynchronously renders selected topics with archive-only resource resolution.</summary>
    public static async Task<PdfDocumentConversionResult> ToPdfDocumentResultAsync(this ChmDocument document, ChmConversionOptions? options = null,
        HtmlToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        ChmConversionOptions configured = options?.Clone() ?? new ChmConversionOptions();
        IReadOnlyList<ChmTopic> topics = document.SelectTopics(configured);
        var diagnostics = document.SourceDiagnostics();
        HtmlToPdfOptions template = pdfOptions?.ClonePdf() ?? new HtmlToPdfOptions();
        if (topics.Count > 1 && template.PdfOptions.TaggedStructureMode != PdfTaggedStructureMode.None) {
            // The PDF owner blocks full rewrites of tagged structure. Preserve topic CSS and page
            // boundaries through an explicitly reported untagged merge instead of bypassing that gate.
            template.PdfOptions.SetTaggedStructureMode(PdfTaggedStructureMode.None);
            diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_PDF_TAGGED_STRUCTURE_OMITTED", "Multi-topic PDF uses untagged pages because the PDF owner cannot preserve a merged structure tree. Select one topic for tagged output.", OfficeConversionLossKind.Omission, "OfficeIMO.Chm.Pdf"));
        }
        PdfStandardEncryptionOptions? encryption = template.PdfOptions.Encryption;
        template.PdfOptions.ClearEncryption();
        var pdfs = new List<PdfDocument>();
        long characters = 0, retainedBytes = 0;
        int nodes = 0;
        foreach (ChmTopic topic in topics) {
            cancellationToken.ThrowIfCancellationRequested();
            string html = topic.ReadHtml(cancellationToken); ChmDocument.ReserveCharacters(ref characters, html.Length, configured);
            HtmlToPdfOptions rendering = template.ClonePdf();
            document.ConfigureRenderOptions(rendering, topic.Path);
            rendering.EmbeddedPackageResourceResolver = rendering.ResourcePolicy.AllowEmbeddedPackageResources ? document.CreateResourceResolver() : null;
            rendering.ResourceResolver = rendering.EmbeddedPackageResourceResolver;
            PdfDocumentConversionResult converted = await document.ParseConversionTopic(html, topic.Path, configured, ref nodes, cancellationToken)
                .ToPdfDocumentResultAsync(rendering, cancellationToken).ConfigureAwait(false);
            // Bound retained merge inputs before accumulating the entire publication.
            retainedBytes += converted.ToBytes(cancellationToken).LongLength;
            ChmDocument.EnforceOutput(retainedBytes, configured);
            pdfs.Add(converted.Value);
            diagnostics.AddRange(converted.FidelityDiagnostics.Select(item => new OfficeConversionFidelityDiagnostic(item.Code, item.Message, item.LossKind, item.Source, topic.Path)));
        }
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocument value = pdfs[0];
        if (pdfs.Count > 1) {
            PdfMergeResult merged = PdfDocument.MergeResult(new PdfMergeOptions {
                Policy = new PdfMergePolicy { Outlines = PdfMergeStructureMode.Combine, NamedDestinations = PdfMergeStructureMode.Combine }
            }, pdfs, cancellationToken);
            value = merged.RequireValue();
            foreach (PdfMergeDecision decision in merged.Report.Decisions.Where(item => item.DroppedCount != 0))
                diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_PDF_MERGE_OMISSION", decision.Action, OfficeConversionLossKind.Omission, "OfficeIMO.Pdf", decision.Structure));
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (encryption != null) {
            PdfSecurityMutationResult encrypted = PdfSecurityEditor.Encrypt(new PdfDocumentConversionResult(value, new PdfConversionReport()).ToBytes(cancellationToken),
                encryption, maximumOutputBytes: configured.MaxOutputBytes, cancellationToken: cancellationToken);
            if (!encrypted.PreservationReport.IsPreserved) throw new InvalidOperationException("CHM PDF encryption did not preserve the merged document.");
            value = encrypted.ToDocument();
        }
        if (topics.Count > 1) diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_PDF_TOPIC_LINKS", "Links between separate help topics are not rebuilt as PDF destinations.", OfficeConversionLossKind.Omission, "OfficeIMO.Chm.Pdf"));
        if (document.Index.Count != 0) diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_PDF_INDEX", "The compiled help keyword index is not recreated in PDF.", OfficeConversionLossKind.Omission, "OfficeIMO.Chm.Pdf"));
        var result = new PdfDocumentConversionResult(value, new PdfConversionReport())
            .WithSourceConversionReport(new ChmConversionReport(topics.Select(topic => topic.Path), diagnostics));
        return result;
    }

    /// <summary>Serializes the selected topic pages as PDF.</summary>
    public static ChmConversionResult<byte[]> ToPdfBytesResult(this ChmDocument document, ChmConversionOptions? options = null,
        HtmlToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        ChmConversionOptions configured = options?.Clone() ?? new ChmConversionOptions();
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(configured, pdfOptions, cancellationToken);
        byte[] bytes = result.ToBytes(cancellationToken); ChmDocument.EnforceOutput(bytes.LongLength, configured);
        return new ChmConversionResult<byte[]>(bytes, new ChmConversionReport(document.SelectTopics(configured).Select(topic => topic.Path), result.FidelityDiagnostics));
    }
}
