namespace OfficeIMO.Pdf;

/// <summary>
/// Zero-dependency helpers for inspecting PDF document metadata and page geometry.
/// </summary>
public static class PdfInspector {
    /// <summary>
    /// Inspects a PDF from a byte array.
    /// </summary>
    public static PdfDocumentInfo Inspect(byte[] pdf, PdfReadOptions? options = null) {
        PdfDocumentProbe probe = Probe(pdf);
        var document = PdfReadDocument.Load(pdf, options);
        return FromReadDocument(document, probe);
    }

    /// <summary>
    /// Inspects a PDF from a file path.
    /// </summary>
    public static PdfDocumentInfo Inspect(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Inspect(File.ReadAllBytes(path), options);
    }

    /// <summary>
    /// Inspects a PDF from the current position of a readable stream.
    /// </summary>
    public static PdfDocumentInfo Inspect(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Inspect(buffer.ToArray(), options);
    }

    /// <summary>
    /// Reports whether OfficeIMO.Pdf can read or safely rewrite a PDF from a byte array.
    /// </summary>
    public static PdfDocumentPreflight Preflight(byte[] pdf, PdfReadOptions? options = null) {
        PdfDocumentProbe probe = Probe(pdf);
        var diagnostics = new List<string>();
        var readBlockers = new List<PdfReadBlocker>();
        var rewriteBlockers = new List<PdfRewriteBlocker>();
        PdfDocumentInfo? info = null;
        PdfReadDocument? readDocument = null;

        if (probe.HeaderVersion is null) {
            AddReadBlocker(PdfReadBlockerKind.MissingHeader, "PDF header was not found.");
        }

        if (probe.HasEncryption) {
            AddReadBlocker(PdfReadBlockerKind.Encryption, "Encrypted PDF files are not supported by OfficeIMO.Pdf yet.");
            AddRewriteBlocker(PdfRewriteBlockerKind.Encryption, "Encrypted PDF files are not supported by OfficeIMO.Pdf yet.");
        }

        bool canRead = diagnostics.Count == 0;
        if (canRead) {
            try {
                readDocument = PdfReadDocument.Load(pdf, options);
                info = FromReadDocument(readDocument, probe);
                if (info.PageCount == 0) {
                    AddReadBlocker(PdfReadBlockerKind.NoPages, "No PDF pages were discovered.");
                    canRead = false;
                }

                var unsupportedContentFilters = GetUnsupportedContentStreamFilters(readDocument);
                if (unsupportedContentFilters.Count > 0) {
                    AddReadBlocker(
                        PdfReadBlockerKind.UnsupportedContentStreamFilter,
                        "PDF page content streams use unsupported filter(s): " + string.Join(", ", unsupportedContentFilters) + ".");
                    canRead = false;
                }
            } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
                AddReadBlocker(PdfReadBlockerKind.ParserUnsupported, "PDF could not be parsed by OfficeIMO.Pdf: " + ex.Message);
                canRead = false;
            }
        }

        if (canRead && readDocument is not null) {
            try {
                ValidateRewriteObjectGraph(pdf, readDocument);
            } catch (Exception ex) when (ex is InvalidOperationException || ex is NotSupportedException || ex is ArgumentException) {
                AddRewriteBlocker(PdfRewriteBlockerKind.InvalidObjectReferences, "PDF object graph is not safe for rewriting by OfficeIMO.Pdf yet: " + ex.Message);
            }
        }

        if (probe.HasSignatures) {
            AddRewriteBlocker(PdfRewriteBlockerKind.Signatures, "Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasForms) {
            AddRewriteBlocker(PdfRewriteBlockerKind.Forms, "PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOutlines && PdfSyntax.HasUnsupportedOutlineRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.Outlines, "PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasPageLabels && PdfSyntax.HasUnsupportedPageLabelRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.PageLabels, "PDF page labels are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasCatalogNameTrees && PdfSyntax.HasUnsupportedCatalogNameTreeRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.CatalogNameTrees, "PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasNamedDestinations && PdfSyntax.HasUnsupportedNamedDestinationRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations, "PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOpenActions && PdfSyntax.HasUnsupportedOpenActionRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.OpenActions, "PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasViewerPreferences && PdfSyntax.HasUnsupportedViewerPreferenceRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.ViewerPreferences, "PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasTaggedContent) {
            AddRewriteBlocker(PdfRewriteBlockerKind.TaggedContent, "PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasXmpMetadata && PdfSyntax.HasUnsupportedXmpMetadataRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.XmpMetadata, "PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasCatalogUri && PdfSyntax.HasUnsupportedCatalogUriRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.CatalogUri, "PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOutputIntents && PdfSyntax.HasUnsupportedOutputIntentRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.OutputIntents, "PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasEmbeddedFiles && PdfSyntax.HasUnsupportedEmbeddedFileRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.EmbeddedFiles, "PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOptionalContent && PdfSyntax.HasUnsupportedOptionalContentRewriteMarkers(pdf)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.OptionalContent, "PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasActiveContent) {
            AddRewriteBlocker(PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        bool canRewrite = canRead && rewriteBlockers.Count == 0;
        return new PdfDocumentPreflight(probe, info, canRead, canRewrite, diagnostics.AsReadOnly(), readBlockers.AsReadOnly(), rewriteBlockers.AsReadOnly());

        void AddReadBlocker(PdfReadBlockerKind kind, string message) {
            AddDiagnostic(message);
            readBlockers.Add(new PdfReadBlocker(kind, message));
        }

        void AddRewriteBlocker(PdfRewriteBlockerKind kind, string message) {
            AddDiagnostic(message);
            rewriteBlockers.Add(new PdfRewriteBlocker(kind, message));
        }

        void AddDiagnostic(string message) {
            if (!diagnostics.Contains(message)) {
                diagnostics.Add(message);
            }
        }
    }

    private static void ValidateRewriteObjectGraph(byte[] pdf, PdfReadDocument document) {
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var catalogState = PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw);
        var collector = new PdfPageExtractor.ObjectCollector(objects);

        for (int i = 0; i < document.Pages.Count; i++) {
            collector.CollectPage(document.Pages[i].ObjectNumber);
        }

        collector.CollectObjectGraph(catalogState.Outlines);
        collector.CollectObjectGraph(catalogState.PageLabels);
        collector.CollectObjectGraph(catalogState.NamedDestinationNameTree);
        collector.CollectObjectGraph(catalogState.XmpMetadata);
        collector.CollectObjectGraph(catalogState.CatalogUri);
        collector.CollectObjectGraph(catalogState.OutputIntents);
        collector.CollectObjectGraph(catalogState.EmbeddedFiles);
        collector.CollectObjectGraph(catalogState.AssociatedFiles);
        collector.CollectObjectGraph(catalogState.OptionalContent);
    }

    /// <summary>
    /// Reports whether OfficeIMO.Pdf can read or safely rewrite a PDF from a file path.
    /// </summary>
    public static PdfDocumentPreflight Preflight(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Preflight(File.ReadAllBytes(path), options);
    }

    /// <summary>
    /// Reports whether OfficeIMO.Pdf can read or safely rewrite a PDF from the current position of a readable stream.
    /// </summary>
    public static PdfDocumentPreflight Preflight(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Preflight(buffer.ToArray(), options);
    }

    private static List<string> GetUnsupportedContentStreamFilters(PdfReadDocument document) {
        var unsupported = new List<string>();
        for (int i = 0; i < document.Pages.Count; i++) {
            foreach (string filterName in document.Pages[i].GetUnsupportedContentStreamFilters()) {
                if (!ContainsFilter(unsupported, filterName)) {
                    unsupported.Add(filterName);
                }
            }
        }

        return unsupported;
    }

    private static bool ContainsFilter(List<string> filters, string filterName) {
        for (int i = 0; i < filters.Count; i++) {
            if (string.Equals(filters[i], filterName, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Reads lightweight PDF markers from a byte array without full document parsing.
    /// </summary>
    public static PdfDocumentProbe Probe(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return new PdfDocumentProbe(
            PdfSyntax.GetHeaderVersion(pdf),
            PdfSyntax.HasEncryptionMarkers(pdf),
            PdfSyntax.HasSignatureMarkers(pdf),
            PdfSyntax.HasFormMarkers(pdf),
            PdfSyntax.HasAnnotationMarkers(pdf),
            PdfSyntax.HasOutlineMarkers(pdf),
            PdfSyntax.HasCatalogViewSettingMarkers(pdf),
            PdfSyntax.HasPageLabelMarkers(pdf),
            PdfSyntax.HasCatalogNameTreeMarkers(pdf),
            PdfSyntax.HasNamedDestinationMarkers(pdf),
            PdfSyntax.HasOpenActionMarkers(pdf),
            PdfSyntax.HasViewerPreferenceMarkers(pdf),
            PdfSyntax.HasTaggedContentMarkers(pdf),
            PdfSyntax.HasXmpMetadataMarkers(pdf),
            PdfSyntax.HasCatalogUriMarkers(pdf),
            PdfSyntax.HasOutputIntentMarkers(pdf),
            PdfSyntax.HasEmbeddedFileMarkers(pdf),
            PdfSyntax.HasOptionalContentMarkers(pdf),
            PdfSyntax.HasActiveContentMarkers(pdf));
    }

    /// <summary>
    /// Reads lightweight PDF markers from a file path without full document parsing.
    /// </summary>
    public static PdfDocumentProbe Probe(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Probe(File.ReadAllBytes(path));
    }

    /// <summary>
    /// Reads lightweight PDF markers from the current position of a readable stream without full document parsing.
    /// </summary>
    public static PdfDocumentProbe Probe(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Probe(buffer.ToArray());
    }

    private static PdfDocumentInfo FromReadDocument(PdfReadDocument document, PdfDocumentProbe probe) {
        var pages = new List<PdfPageInfo>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            var (width, height) = document.Pages[i].GetPageSize();
            int rotation = document.Pages[i].GetRotationDegrees();
            int pageNumber = i + 1;
            var pageLinks = document.Pages[i].GetLinkAnnotations();
            var links = new List<PdfLinkAnnotation>(pageLinks.Count);
            for (int j = 0; j < pageLinks.Count; j++) {
                links.Add(pageLinks[j].WithPageNumber(pageNumber));
            }

            pages.Add(new PdfPageInfo(pageNumber, width, height, rotation, links));
        }

        return new PdfDocumentInfo(pages.AsReadOnly(), document.Metadata, document.Outlines, document.PageLabels, document.NamedDestinations, document.OpenAction, document.ViewerPreferences, document.FormFields, document.AcroFormNeedAppearances, document.AcroFormSignatureFlags, probe.HeaderVersion, document.CatalogPageMode, document.CatalogPageLayout, document.CatalogVersion, document.CatalogLanguage, probe.HasSignatures, probe.HasForms, probe.HasAnnotations, probe.HasOutlines, probe.HasCatalogViewSettings, probe.HasPageLabels, probe.HasCatalogNameTrees, probe.HasNamedDestinations, probe.HasOpenActions, probe.HasViewerPreferences, probe.HasTaggedContent, probe.HasXmpMetadata, probe.HasCatalogUri, probe.HasOutputIntents, probe.HasEmbeddedFiles, probe.HasOptionalContent, probe.HasActiveContent);
    }
}
