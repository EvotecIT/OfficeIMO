namespace OfficeIMO.Pdf;

/// <summary>
/// Zero-dependency helpers for inspecting PDF document metadata and page geometry.
/// </summary>
internal static class PdfInspector {
    /// <summary>
    /// Inspects a PDF from a byte array.
    /// </summary>
    public static PdfDocumentInfo Inspect(byte[] pdf, PdfReadOptions? options = null) {
        PdfDocumentProbe probe = Probe(pdf, options);
        var document = PdfReadDocument.Open(pdf, options);
        return FromReadDocument(document, probe);
    }

    internal static PdfDocumentInfo Inspect(byte[] pdf, PdfReadDocument document) =>
        FromReadDocument(document, Probe(pdf, document));

    /// <summary>
    /// Inspects selected source page ranges from a PDF byte array, preserving caller order and overlaps.
    /// </summary>
    public static PdfDocumentInfo InspectPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return InspectPageRanges(pdf, null, pageRanges);
    }

    /// <summary>
    /// Inspects selected source page ranges from a PDF byte array, preserving caller order and overlaps.
    /// </summary>
    public static PdfDocumentInfo InspectPageRanges(byte[] pdf, PdfReadOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfDocumentProbe probe = Probe(pdf, options);
        var document = PdfReadDocument.Open(pdf, options);
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
        return FromReadDocument(document, probe, pageNumbers);
    }

    /// <summary>
    /// Inspects a PDF from a file path.
    /// </summary>
    public static PdfDocumentInfo Inspect(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Inspect(File.ReadAllBytes(path), options);
    }

    /// <summary>
    /// Inspects selected source page ranges from a PDF file path, preserving caller order and overlaps.
    /// </summary>
    public static PdfDocumentInfo InspectPageRanges(string path, params PdfPageRange[] pageRanges) {
        return InspectPageRanges(path, null, pageRanges);
    }

    /// <summary>
    /// Inspects selected source page ranges from a PDF file path, preserving caller order and overlaps.
    /// </summary>
    public static PdfDocumentInfo InspectPageRanges(string path, PdfReadOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return InspectPageRanges(File.ReadAllBytes(path), options, pageRanges);
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
    /// Inspects selected source page ranges from the current position of a readable stream, preserving caller order and overlaps.
    /// </summary>
    public static PdfDocumentInfo InspectPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return InspectPageRanges(stream, null, pageRanges);
    }

    /// <summary>
    /// Inspects selected source page ranges from the current position of a readable stream, preserving caller order and overlaps.
    /// </summary>
    public static PdfDocumentInfo InspectPageRanges(Stream stream, PdfReadOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return InspectPageRanges(buffer.ToArray(), options, pageRanges);
    }

    /// <summary>
    /// Reports whether OfficeIMO.Pdf can read or safely rewrite a PDF from a byte array.
    /// </summary>
    public static PdfDocumentPreflight Preflight(byte[] pdf, PdfReadOptions? options = null) {
        return PreflightCore(pdf, options, readDocumentFactory: null);
    }

    internal static PdfDocumentPreflight Preflight(
        byte[] pdf,
        PdfReadOptions options,
        Func<PdfReadDocument> readDocumentFactory) {
        Guard.NotNull(readDocumentFactory, nameof(readDocumentFactory));
        return PreflightCore(pdf, options, readDocumentFactory);
    }

    private static PdfDocumentPreflight PreflightCore(
        byte[] pdf,
        PdfReadOptions? options,
        Func<PdfReadDocument>? readDocumentFactory) {
        PdfDocumentProbe probe = Probe(pdf, options);
        var diagnostics = new List<string>();
        var readBlockers = new List<PdfReadBlocker>();
        var rewriteBlockers = new List<PdfRewriteBlocker>();
        PdfDocumentInfo? info = null;
        PdfReadDocument? readDocument = null;

        if (probe.HeaderVersion is null) {
            AddReadBlocker(PdfReadBlockerKind.MissingHeader, "PDF header was not found.");
        }

        if (probe.HasEncryption) {
            AddRewriteBlocker(PdfRewriteBlockerKind.Encryption, "General encrypted-document rewrites remain blocked; the dedicated owner-authorized security editor can decrypt or re-encrypt supported unsigned PDFs.");
        }

        bool canRead = readBlockers.Count == 0;
        if (canRead) {
            try {
                readDocument = readDocumentFactory?.Invoke() ?? PdfReadDocument.Open(pdf, options);
                probe = Probe(pdf, readDocument);
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
            } catch (PdfPasswordRequiredException ex) {
                AddReadBlocker(PdfReadBlockerKind.Encryption, ex.Message);
                canRead = false;
            } catch (PdfInvalidPasswordException ex) {
                AddReadBlocker(PdfReadBlockerKind.Encryption, ex.Message);
                canRead = false;
            } catch (PdfUnsupportedEncryptionException ex) {
                AddReadBlocker(PdfReadBlockerKind.Encryption, ex.Message);
                canRead = false;
            } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
                AddReadBlocker(PdfReadBlockerKind.ParserUnsupported, "PDF could not be parsed by OfficeIMO.Pdf: " + ex.Message);
                canRead = false;
            }
        }

        if (canRead && readDocument is not null && !probe.HasEncryption) {
            try {
                ValidateRewriteObjectGraph(readDocument);
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

        if (probe.HasOutlines && PdfSyntax.HasUnsupportedOutlineRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.Outlines, "PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasPageLabels && PdfSyntax.HasUnsupportedPageLabelRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.PageLabels, "PDF page labels are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasCatalogNameTrees && PdfSyntax.HasUnsupportedCatalogNameTreeRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.CatalogNameTrees, "PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasNamedDestinations && PdfSyntax.HasUnsupportedNamedDestinationRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations, "PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOpenActions && PdfSyntax.HasUnsupportedOpenActionRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.OpenActions, "PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasViewerPreferences && PdfSyntax.HasUnsupportedViewerPreferenceRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.ViewerPreferences, "PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasTaggedContent) {
            AddRewriteBlocker(PdfRewriteBlockerKind.TaggedContent, "PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasXmpMetadata && PdfSyntax.HasUnsupportedXmpMetadataRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.XmpMetadata, "PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasCatalogUri && PdfSyntax.HasUnsupportedCatalogUriRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.CatalogUri, "PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOutputIntents && PdfSyntax.HasUnsupportedOutputIntentRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.OutputIntents, "PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasEmbeddedFiles && PdfSyntax.HasUnsupportedEmbeddedFileRewriteMarkers(pdf, options)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.EmbeddedFiles, "PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOptionalContent && PdfSyntax.HasUnsupportedOptionalContentRewriteMarkers(pdf, options)) {
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

    private static void ValidateRewriteObjectGraph(PdfReadDocument document) {
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        string trailerRaw = document.TrailerRaw;
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
        PdfDocumentSource source = PdfDocumentSource.FromPath(path, options);
        return Preflight(source.Bytes, source.Options);
    }

    /// <summary>
    /// Reports whether OfficeIMO.Pdf can read or safely rewrite a PDF from the current position of a readable stream.
    /// </summary>
    public static PdfDocumentPreflight Preflight(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        PdfReadOptions effectiveOptions = PdfReadOptions.Resolve(options);
        long limit = effectiveOptions.Limits.MaxInputBytes;
        if (stream.CanSeek) {
            long remaining = stream.Length - stream.Position;
            if (remaining > limit) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limit, remaining);
            }
        }

        using var buffer = new MemoryStream();
        var chunk = new byte[81920];
        int read;
        while ((read = stream.Read(chunk, 0, chunk.Length)) > 0) {
            long nextLength = buffer.Length + read;
            if (nextLength > limit) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limit, nextLength);
            }

            buffer.Write(chunk, 0, read);
        }

        return Preflight(buffer.ToArray(), effectiveOptions);
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
    public static PdfDocumentProbe Probe(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, options);
        try {
            var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, options);
            return Probe(pdf, security, objects, trailerRaw);
        } catch (Exception ex) when (
            ex is not PdfEncryptionException &&
            ex is not OutOfMemoryException &&
            ex is not StackOverflowException) {
            return ProbeFromRawBytes(pdf, security, options);
        } catch (PdfEncryptionException) when (options?.Password is null) {
            return ProbeFromRawBytes(pdf, security, options);
        }
    }

    internal static PdfDocumentProbe Probe(byte[] pdf, PdfReadDocument document) =>
        Probe(pdf, document.Security, document.Objects, document.TrailerRaw);

    private static PdfDocumentProbe Probe(
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        Dictionary<int, PdfIndirectObject> objects,
        string trailerRaw) {
        string text = PdfEncoding.Latin1GetString(pdf);
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        bool Has(params string[] names) =>
            PdfSyntax.ContainsAnyPdfName(text, names) ||
            PdfSyntax.ContainsAnyParsedPdfName(objects, names);

        return new PdfDocumentProbe(
            PdfSyntax.GetHeaderVersion(pdf),
            security.HasEncryption,
            Has("ByteRange", "SigFlags", "Sig"),
            Has("AcroForm", "Fields", "FT", "XFA"),
            Has("Annots", "Annot"),
            Has("Outlines", "UseOutlines"),
            Has("PageMode", "PageLayout"),
            Has("PageLabels"),
            Has("Names"),
            Has("Dests"),
            Has("OpenAction"),
            Has("ViewerPreferences"),
            Has("MarkInfo", "StructTreeRoot", "ParentTree", "StructElem"),
            Has("Metadata"),
            catalog?.Items.ContainsKey("URI") == true,
            Has("OutputIntents", "OutputIntent"),
            Has("EmbeddedFiles", "Filespec", "EmbeddedFile", "AF"),
            Has("OCProperties", "OCGs", "OCG", "OCMD"),
            Has("JavaScript", "JS", "AA", "Launch", "SubmitForm", "RichMedia"),
            security);
    }

    private static PdfDocumentProbe ProbeFromRawBytes(
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        PdfReadOptions? options) =>
        new PdfDocumentProbe(
            PdfSyntax.GetHeaderVersion(pdf),
            security.HasEncryption,
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
            PdfSyntax.HasCatalogUriMarkers(pdf, options),
            PdfSyntax.HasOutputIntentMarkers(pdf),
            PdfSyntax.HasEmbeddedFileMarkers(pdf),
            PdfSyntax.HasOptionalContentMarkers(pdf),
            PdfSyntax.HasActiveContentMarkers(pdf),
            security);

    /// <summary>
    /// Reads lightweight PDF markers from a file path without full document parsing.
    /// </summary>
    public static PdfDocumentProbe Probe(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Probe(File.ReadAllBytes(path), options);
    }

    /// <summary>
    /// Reads lightweight PDF markers from the current position of a readable stream without full document parsing.
    /// </summary>
    public static PdfDocumentProbe Probe(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Probe(buffer.ToArray(), options);
    }

    internal static PdfDocumentInfo FromReadDocument(PdfReadDocument document, PdfDocumentProbe probe, int[]? pageNumbers = null) {
        pageNumbers ??= PdfPageRangeObjectFilter.GetAllPageNumbers(document.Pages.Count);
        bool useDocumentWideObjects = PdfPageRangeObjectFilter.ShouldUseDocumentWideObjects(document.Pages.Count, pageNumbers);
        IReadOnlyList<PdfFormField> formFields = useDocumentWideObjects
            ? document.FormFields
            : PdfPageRangeObjectFilter.FilterFormFieldsByPageNumbers(document.FormFields, pageNumbers, preservePageDuplicates: true);
        IReadOnlyList<PdfOutlineItem> outlines = useDocumentWideObjects
            ? document.Outlines
            : PdfPageRangeObjectFilter.FilterOutlinesByPageNumbers(document.Outlines, pageNumbers);
        IReadOnlyList<PdfPageLabel> pageLabels = useDocumentWideObjects
            ? document.PageLabels
            : PdfPageRangeObjectFilter.FilterPageLabelsByPageNumbers(document.PageLabels, pageNumbers);
        IReadOnlyList<PdfNamedDestination> namedDestinations = useDocumentWideObjects
            ? document.NamedDestinations
            : PdfPageRangeObjectFilter.FilterNamedDestinationsByPageNumbers(document.NamedDestinations, pageNumbers);
        IReadOnlyList<PdfCatalogAction> catalogActions = useDocumentWideObjects
            ? document.CatalogActions
            : Array.Empty<PdfCatalogAction>();
        IReadOnlyList<PdfAttachmentInfo> attachments = useDocumentWideObjects
            ? document.Attachments
            : Array.Empty<PdfAttachmentInfo>();
        IReadOnlyList<PdfOutputIntentInfo> outputIntents = useDocumentWideObjects
            ? document.OutputIntents
            : Array.Empty<PdfOutputIntentInfo>();
        PdfXmpMetadataInfo? xmpMetadata = useDocumentWideObjects
            ? document.XmpMetadata
            : null;
        PdfTaggedContentInfo? taggedContent = useDocumentWideObjects
            ? document.TaggedContent
            : null;
        PdfOptionalContentProperties? optionalContent = useDocumentWideObjects
            ? document.OptionalContent
            : null;
        PdfDocumentOpenAction? openAction = useDocumentWideObjects
            ? document.OpenAction
            : PdfPageRangeObjectFilter.FilterOpenActionByPageNumbers(document.OpenAction, pageNumbers);

        var pages = new List<PdfPageInfo>(pageNumbers.Length);
        var widgetsByPage = BuildFormWidgetsByPage(document.FormFields);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            PdfReadPage page = document.Pages[pageNumber - 1];
            PdfPageGeometry geometry = page.GetGeometry();
            var (width, height) = page.GetPageSize();
            int rotation = page.GetRotationDegrees();
            var pageLinks = page.GetLinkAnnotations();
            var links = new List<PdfLinkAnnotation>(pageLinks.Count);
            for (int j = 0; j < pageLinks.Count; j++) {
                PdfLinkAnnotation link = pageLinks[j].WithPageNumber(pageNumber);
                if (link.DestinationPageObjectNumber.HasValue) {
                    link = link.WithDestinationPageNumber(document.GetPageNumberForObject(link.DestinationPageObjectNumber.Value));
                }

                links.Add(link);
            }

            var pageAnnotations = page.GetAnnotations();
            var annotations = new List<PdfAnnotation>(pageAnnotations.Count);
            for (int j = 0; j < pageAnnotations.Count; j++) {
                annotations.Add(pageAnnotations[j].WithPageNumber(pageNumber));
            }

            var pageActions = page.GetPageActions();
            var actions = new List<PdfPageAction>(pageActions.Count);
            for (int j = 0; j < pageActions.Count; j++) {
                actions.Add(pageActions[j].WithPageNumber(pageNumber));
            }

            widgetsByPage.TryGetValue(pageNumber, out IReadOnlyList<PdfFormWidget>? formWidgets);
            pages.Add(new PdfPageInfo(pageNumber, width, height, rotation, geometry, links, formWidgets, annotations, actions));
        }

        return new PdfDocumentInfo(pages.AsReadOnly(), document.Metadata, outlines, pageLabels, namedDestinations, catalogActions, attachments, outputIntents, xmpMetadata, taggedContent, optionalContent, openAction, document.ViewerPreferences, formFields, document.AcroFormDefaultAppearance, document.AcroFormQuadding, document.AcroFormXfa, document.AcroFormNeedAppearances, document.AcroFormSignatureFlags, document.Security, probe.HeaderVersion, document.CatalogPageMode, document.CatalogPageLayout, document.CatalogVersion, document.CatalogLanguage, document.Security.HasSignatures || probe.HasSignatures, probe.HasForms || document.AcroFormXfa is not null, probe.HasAnnotations, probe.HasOutlines, probe.HasCatalogViewSettings, probe.HasPageLabels, probe.HasCatalogNameTrees, probe.HasNamedDestinations, probe.HasOpenActions, probe.HasViewerPreferences, probe.HasTaggedContent, probe.HasXmpMetadata, probe.HasCatalogUri, probe.HasOutputIntents, probe.HasEmbeddedFiles, probe.HasOptionalContent, probe.HasActiveContent);
    }

    private static Dictionary<int, IReadOnlyList<PdfFormWidget>> BuildFormWidgetsByPage(IReadOnlyList<PdfFormField> fields) {
        var grouped = new Dictionary<int, List<PdfFormWidget>>();
        for (int i = 0; i < fields.Count; i++) {
            IReadOnlyList<PdfFormWidget> widgets = fields[i].Widgets;
            for (int j = 0; j < widgets.Count; j++) {
                PdfFormWidget widget = widgets[j];
                if (!widget.PageNumber.HasValue) {
                    continue;
                }

                if (!grouped.TryGetValue(widget.PageNumber.Value, out List<PdfFormWidget>? pageWidgets)) {
                    pageWidgets = new List<PdfFormWidget>();
                    grouped.Add(widget.PageNumber.Value, pageWidgets);
                }

                pageWidgets.Add(widget);
            }
        }

        var result = new Dictionary<int, IReadOnlyList<PdfFormWidget>>();
        foreach (var item in grouped) {
            result.Add(item.Key, item.Value.AsReadOnly());
        }

        return result;
    }
}
