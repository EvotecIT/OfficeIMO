using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Zero-dependency helpers for inspecting PDF document metadata and page geometry.
/// </summary>
internal static class PdfInspector {
    private enum ProbeMarker {
        Signatures, Forms, Annotations, Outlines, CatalogView, PageLabels, Names,
        Destinations, OpenActions, ViewerPreferences, TaggedContent, Metadata,
        OutputIntents, EmbeddedFiles, Uri, OptionalContent, ActiveContent
    }

    // One marker catalog serves parsed inspection and the conservative raw fallback.
    private static readonly string[][] ProbeMarkerNames = {
        new[] { "ByteRange", "SigFlags", "Sig" },
        new[] { "AcroForm", "Fields", "FT", "XFA" },
        new[] { "Annots", "Annot" },
        new[] { "Outlines", "UseOutlines" },
        new[] { "PageMode", "PageLayout" },
        new[] { "PageLabels" },
        new[] { "Names" },
        new[] { "Dests" },
        new[] { "OpenAction" },
        new[] { "ViewerPreferences" },
        new[] { "MarkInfo", "StructTreeRoot", "ParentTree", "StructElem" },
        new[] { "Metadata" },
        new[] { "OutputIntents", "OutputIntent" },
        new[] { "EmbeddedFiles", "Filespec", "EmbeddedFile", "AF" },
        new[] { "URI" },
        new[] { "OCProperties", "OCGs", "OCG", "OCMD" },
        PdfActiveContentPolicy.MarkerNames
    };

    private static readonly HashSet<string> ParsedProbeMarkerNames = CreateParsedProbeMarkerNames();

    private static HashSet<string> CreateParsedProbeMarkerNames() {
        var names = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index <= (int)ProbeMarker.EmbeddedFiles; index++) {
            foreach (string name in ProbeMarkerNames[index]) names.Add(name);
        }
        return names;
    }

    /// <summary>
    /// Inspects a PDF from a byte array.
    /// </summary>
    public static PdfDocumentInfo Inspect(byte[] pdf, PdfLoadOptions? options = null) {
        PdfDocumentProbe probe = Probe(pdf, options);
        var document = PdfReadDocument.Open(pdf, options);
        return FromReadDocument(document, probe);
    }

    internal static PdfDocumentInfo Inspect(byte[] pdf, PdfReadDocument document) =>
        FromReadDocument(document, Probe(pdf, document));

    internal static PdfDocumentInfo Inspect(
        byte[] pdf,
        PdfReadDocument document,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return FromReadDocument(
            document,
            Probe(pdf, document, cancellationToken),
            pageNumbers: null,
            cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Inspects selected source page ranges from a PDF byte array, preserving caller order and overlaps.
    /// </summary>
    public static PdfDocumentInfo InspectPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return InspectPageRanges(pdf, null, pageRanges);
    }

    /// <summary>
    /// Inspects selected source page ranges from a PDF byte array, preserving caller order and overlaps.
    /// </summary>
    public static PdfDocumentInfo InspectPageRanges(byte[] pdf, PdfLoadOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfDocumentProbe probe = Probe(pdf, options);
        var document = PdfReadDocument.Open(pdf, options);
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
        return FromReadDocument(document, probe, pageNumbers);
    }

    /// <summary>
    /// Inspects a PDF from a file path.
    /// </summary>
    public static PdfDocumentInfo Inspect(string path, PdfLoadOptions? options = null) {
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
    public static PdfDocumentInfo InspectPageRanges(string path, PdfLoadOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return InspectPageRanges(File.ReadAllBytes(path), options, pageRanges);
    }

    /// <summary>
    /// Inspects a PDF from the current position of a readable stream.
    /// </summary>
    public static PdfDocumentInfo Inspect(Stream stream, PdfLoadOptions? options = null) {
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
    public static PdfDocumentInfo InspectPageRanges(Stream stream, PdfLoadOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return InspectPageRanges(buffer.ToArray(), options, pageRanges);
    }

    /// <summary>
    /// Reports whether OfficeIMO.Pdf can read or safely rewrite a PDF from a byte array.
    /// </summary>
    public static PdfDocumentPreflight Preflight(byte[] pdf, PdfLoadOptions? options = null) {
        return PreflightCore(pdf, options, readDocumentFactory: null);
    }

    internal static PdfDocumentPreflight Preflight(
        byte[] pdf,
        PdfLoadOptions? options,
        CancellationToken cancellationToken) {
        return PreflightCore(pdf, options, readDocumentFactory: null, cancellationToken);
    }

    internal static PdfDocumentPreflight Preflight(
        byte[] pdf,
        PdfLoadOptions options,
        Func<PdfReadDocument> readDocumentFactory) {
        Guard.NotNull(readDocumentFactory, nameof(readDocumentFactory));
        return PreflightCore(pdf, options, readDocumentFactory);
    }

    internal static PdfDocumentPreflight Preflight(
        byte[] pdf,
        PdfLoadOptions options,
        Func<PdfReadDocument> readDocumentFactory,
        CancellationToken cancellationToken) {
        Guard.NotNull(readDocumentFactory, nameof(readDocumentFactory));
        return PreflightCore(pdf, options, readDocumentFactory, cancellationToken);
    }

    private static PdfDocumentPreflight PreflightCore(
        byte[] pdf,
        PdfLoadOptions? options,
        Func<PdfReadDocument>? readDocumentFactory,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        PdfReadDocument? readDocument = null;
        Exception? readDocumentException = null;
        PdfDocumentProbe probe;
        bool probeFromReadDocument = false;
        try {
            readDocument = readDocumentFactory is null
                ? PdfReadDocument.Open(pdf, effectiveOptions, cancellationToken)
                : readDocumentFactory();
            probe = Probe(pdf, readDocument, cancellationToken);
            probeFromReadDocument = true;
        } catch (Exception ex) when (ex is not OperationCanceledException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            readDocumentException = ex;
            probe = Probe(pdf, effectiveOptions, cancellationToken);
        }

        var diagnostics = new List<string>();
        var readBlockers = new List<PdfReadBlocker>();
        var rewriteBlockers = new List<PdfRewriteBlocker>();
        PdfDocumentInfo? info = null;

        if (probe.HeaderVersion is null) {
            AddReadBlocker(PdfReadBlockerKind.MissingHeader, "PDF header was not found.");
        }

        if (probe.HasEncryption) {
            AddRewriteBlocker(PdfRewriteBlockerKind.Encryption, "Encrypted input requires operation-specific planning. Authenticated unsigned PDFs support proven page, metadata, sanitization, and simple form rewrites when the required permissions are authorized; security changes require owner authorization.");
        }

        bool canRead = readBlockers.Count == 0;
        if (canRead) {
            try {
                if (readDocumentException is not null) {
                    throw readDocumentException;
                }

                readDocument ??= PdfReadDocument.Open(pdf, effectiveOptions, cancellationToken);
                if (!probeFromReadDocument) {
                    probe = Probe(pdf, readDocument, cancellationToken);
                }
                info = FromReadDocument(readDocument, probe, cancellationToken: cancellationToken);
                if (info.PageCount == 0) {
                    AddReadBlocker(PdfReadBlockerKind.NoPages, "No PDF pages were discovered.");
                    canRead = false;
                }

                var unsupportedContentFilters = GetUnsupportedContentStreamFilters(readDocument, cancellationToken);
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
            } catch (Exception ex) when (ex is not OperationCanceledException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
                AddReadBlocker(PdfReadBlockerKind.ParserUnsupported, "PDF could not be parsed by OfficeIMO.Pdf: " + ex.Message);
                canRead = false;
            }
        }

        if (readDocument?.RepairReport.HasUnreadableObjects == true)
            AddRewriteBlocker(PdfRewriteBlockerKind.IncompleteObjectGraph,
                "PDF contains unreadable indirect objects; mutation cannot prove preservation or signature safety.");

        if (canRead && readDocument is not null && !probe.HasEncryption) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                ValidateRewriteObjectGraph(readDocument, cancellationToken);
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

        var rewriteMarkerSource = new PdfRewriteMarkerSource(pdf, effectiveOptions, readDocument);

        if (probe.HasOutlines && PdfSyntax.HasUnsupportedOutlineRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.Outlines, "PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasPageLabels && PdfSyntax.HasUnsupportedPageLabelRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.PageLabels, "PDF page labels are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasCatalogNameTrees && PdfSyntax.HasUnsupportedCatalogNameTreeRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.CatalogNameTrees, "PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasNamedDestinations && PdfSyntax.HasUnsupportedNamedDestinationRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations, "PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOpenActions && PdfSyntax.HasUnsupportedOpenActionRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.OpenActions, "PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasViewerPreferences && PdfSyntax.HasUnsupportedViewerPreferenceRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.ViewerPreferences, "PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasTaggedContent) {
            AddRewriteBlocker(PdfRewriteBlockerKind.TaggedContent, "PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasXmpMetadata && PdfSyntax.HasUnsupportedXmpMetadataRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.XmpMetadata, "PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasCatalogUri && PdfSyntax.HasUnsupportedCatalogUriRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.CatalogUri, "PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOutputIntents && PdfSyntax.HasUnsupportedOutputIntentRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.OutputIntents, "PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasEmbeddedFiles && PdfSyntax.HasUnsupportedEmbeddedFileRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.EmbeddedFiles, "PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasOptionalContent && PdfSyntax.HasUnsupportedOptionalContentRewriteMarkers(pdf, options, rewriteMarkerSource)) {
            AddRewriteBlocker(PdfRewriteBlockerKind.OptionalContent, "PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (probe.HasActiveContent) {
            AddRewriteBlocker(PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        cancellationToken.ThrowIfCancellationRequested();
        bool canRewrite = canRead && rewriteBlockers.Count == 0;
        return new PdfDocumentPreflight(probe, info, canRead, canRewrite, diagnostics.AsReadOnly(), readBlockers.AsReadOnly(), rewriteBlockers.AsReadOnly(), effectiveOptions.PermissionPolicy);

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

    private static void ValidateRewriteObjectGraph(
        PdfReadDocument document,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        string trailerRaw = document.TrailerRaw;
        var catalogState = PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw);
        var collector = new PdfPageExtractor.ObjectCollector(objects, cancellationToken: cancellationToken);

        for (int i = 0; i < document.Pages.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
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
    public static PdfDocumentPreflight Preflight(string path, PdfLoadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        PdfDocumentSource source = PdfDocumentSource.FromPath(path, options);
        return Preflight(source.Bytes, source.Options);
    }

    /// <summary>
    /// Reports whether OfficeIMO.Pdf can read or safely rewrite a PDF from the current position of a readable stream.
    /// </summary>
    public static PdfDocumentPreflight Preflight(Stream stream, PdfLoadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
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

    private static List<string> GetUnsupportedContentStreamFilters(
        PdfReadDocument document,
        CancellationToken cancellationToken) {
        var unsupported = new List<string>();
        for (int i = 0; i < document.Pages.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (string filterName in document.Pages[i].GetUnsupportedContentStreamFilters()) {
                cancellationToken.ThrowIfCancellationRequested();
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
    public static PdfDocumentProbe Probe(byte[] pdf, PdfLoadOptions? options = null) {
        return Probe(pdf, options, CancellationToken.None);
    }

    internal static PdfDocumentProbe Probe(
        byte[] pdf,
        PdfLoadOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(pdf, nameof(pdf));
        cancellationToken.ThrowIfCancellationRequested();

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(
            pdf,
            options,
            cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        try {
            var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, options, out PdfRepairReport repairReport, out _, cancellationToken);
            return Probe(pdf, security, objects, trailerRaw, repairReport, cancellationToken);
        } catch (Exception ex) when (
            ex is not PdfEncryptionException &&
            ex is not OperationCanceledException &&
            ex is not OutOfMemoryException &&
            ex is not StackOverflowException) {
            cancellationToken.ThrowIfCancellationRequested();
            return ProbeFromRawBytes(pdf, security, cancellationToken);
        } catch (PdfEncryptionException) when (options?.Password is null) {
            cancellationToken.ThrowIfCancellationRequested();
            return ProbeFromRawBytes(pdf, security, cancellationToken);
        }
    }

    internal static PdfDocumentProbe Probe(byte[] pdf, PdfReadDocument document) =>
        Probe(pdf, document.Security, document.Objects, document.TrailerRaw, document.RepairReport);

    internal static PdfDocumentProbe Probe(
        byte[] pdf,
        PdfReadDocument document,
        CancellationToken cancellationToken) =>
        Probe(pdf, document.Security, document.Objects, document.TrailerRaw, document.RepairReport, cancellationToken);

    private static PdfDocumentProbe Probe(
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        Dictionary<int, PdfIndirectObject> objects,
        string trailerRaw,
        PdfRepairReport repairReport,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        // A single parsed walk replaces a separate full graph scan for every feature.
        HashSet<string> presentNames = PdfSyntax.CollectParsedPdfNames(objects, ParsedProbeMarkerNames, cancellationToken);
        string? rawFallback = repairReport.HasIncompleteObjectCoverage ? PdfEncoding.Latin1GetString(pdf) : null;
        bool Has(ProbeMarker marker) {
            cancellationToken.ThrowIfCancellationRequested();
            // Parsed dictionaries are authoritative here. Stream bytes and string values
            // can contain marker-shaped text, including random encrypted payload bytes.
            string[] names = ProbeMarkerNames[(int)marker];
            foreach (string name in names) {
                if (presentNames.Contains(name)) return true;
            }
            bool found = rawFallback != null && PdfSyntax.ContainsAnyPdfName(rawFallback, cancellationToken, names);
            cancellationToken.ThrowIfCancellationRequested();
            return found;
        }
        bool HasReachable(ProbeMarker marker) {
            cancellationToken.ThrowIfCancellationRequested();
            bool found = catalog != null &&
                PdfSyntax.ContainsAnyReachableParsedPdfName(catalog, objects, ProbeMarkerNames[(int)marker]);
            cancellationToken.ThrowIfCancellationRequested();
            return found;
        }

        return new PdfDocumentProbe(
            PdfSyntax.GetHeaderVersion(pdf),
            security.HasEncryption,
            Has(ProbeMarker.Signatures),
            Has(ProbeMarker.Forms),
            Has(ProbeMarker.Annotations),
            Has(ProbeMarker.Outlines),
            Has(ProbeMarker.CatalogView),
            Has(ProbeMarker.PageLabels),
            Has(ProbeMarker.Names),
            Has(ProbeMarker.Destinations),
            Has(ProbeMarker.OpenActions),
            Has(ProbeMarker.ViewerPreferences),
            Has(ProbeMarker.TaggedContent),
            Has(ProbeMarker.Metadata),
            catalog?.Items.ContainsKey("URI") == true,
            Has(ProbeMarker.OutputIntents),
            Has(ProbeMarker.EmbeddedFiles),
            HasReachable(ProbeMarker.OptionalContent),
            HasReachable(ProbeMarker.ActiveContent),
            security);
    }

    private static PdfDocumentProbe ProbeFromRawBytes(
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string text = PdfEncoding.Latin1GetString(pdf);
        cancellationToken.ThrowIfCancellationRequested();
        bool Has(ProbeMarker marker) => PdfSyntax.ContainsAnyPdfName(text, cancellationToken, ProbeMarkerNames[(int)marker]);

        return new PdfDocumentProbe(
            PdfSyntax.GetHeaderVersion(pdf),
            security.HasEncryption,
            Has(ProbeMarker.Signatures),
            Has(ProbeMarker.Forms),
            Has(ProbeMarker.Annotations),
            Has(ProbeMarker.Outlines),
            Has(ProbeMarker.CatalogView),
            Has(ProbeMarker.PageLabels),
            Has(ProbeMarker.Names),
            Has(ProbeMarker.Destinations),
            Has(ProbeMarker.OpenActions),
            Has(ProbeMarker.ViewerPreferences),
            Has(ProbeMarker.TaggedContent),
            Has(ProbeMarker.Metadata),
            Has(ProbeMarker.Uri),
            Has(ProbeMarker.OutputIntents),
            Has(ProbeMarker.EmbeddedFiles),
            Has(ProbeMarker.OptionalContent),
            Has(ProbeMarker.ActiveContent),
            security);
    }

    /// <summary>
    /// Reads lightweight PDF markers from a file path without full document parsing.
    /// </summary>
    public static PdfDocumentProbe Probe(string path, PdfLoadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Probe(File.ReadAllBytes(path), options);
    }

    /// <summary>
    /// Reads lightweight PDF markers from the current position of a readable stream without full document parsing.
    /// </summary>
    public static PdfDocumentProbe Probe(Stream stream, PdfLoadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Probe(buffer.ToArray(), options);
    }

    internal static PdfDocumentInfo FromReadDocument(
        PdfReadDocument document,
        PdfDocumentProbe probe,
        int[]? pageNumbers = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        pageNumbers ??= PdfPageRangeObjectFilter.GetAllPageNumbers(document.Pages.Count);
        bool useDocumentWideObjects = PdfPageRangeObjectFilter.ShouldUseDocumentWideObjects(document.Pages.Count, pageNumbers);
        IReadOnlyList<PdfFormField> formFields = useDocumentWideObjects
            ? document.UncheckedFormFields
            : PdfPageRangeObjectFilter.FilterFormFieldsByPageNumbers(document.UncheckedFormFields, pageNumbers, preservePageDuplicates: true);
        IReadOnlyList<PdfOutlineItem> outlines = useDocumentWideObjects
            ? document.UncheckedOutlines
            : PdfPageRangeObjectFilter.FilterOutlinesByPageNumbers(document.UncheckedOutlines, pageNumbers);
        IReadOnlyList<PdfPageLabel> pageLabels = useDocumentWideObjects
            ? document.UncheckedPageLabels
            : PdfPageRangeObjectFilter.FilterPageLabelsByPageNumbers(document.UncheckedPageLabels, pageNumbers);
        IReadOnlyList<PdfNamedDestination> namedDestinations = useDocumentWideObjects
            ? document.UncheckedNamedDestinations
            : PdfPageRangeObjectFilter.FilterNamedDestinationsByPageNumbers(document.UncheckedNamedDestinations, pageNumbers);
        IReadOnlyList<PdfCatalogAction> catalogActions = useDocumentWideObjects
            ? document.UncheckedCatalogActions
            : Array.Empty<PdfCatalogAction>();
        IReadOnlyList<PdfAttachmentInfo> attachments = useDocumentWideObjects
            ? document.UncheckedAttachments
            : Array.Empty<PdfAttachmentInfo>();
        IReadOnlyList<PdfOutputIntentInfo> outputIntents = useDocumentWideObjects
            ? document.UncheckedOutputIntents
            : Array.Empty<PdfOutputIntentInfo>();
        bool outputIntentsAreComplete = useDocumentWideObjects && document.UncheckedOutputIntentsAreComplete;
        PdfXmpMetadataInfo? xmpMetadata = useDocumentWideObjects
            ? document.UncheckedXmpMetadata
            : null;
        PdfTaggedContentInfo? taggedContent = useDocumentWideObjects
            ? document.UncheckedTaggedContent
            : null;
        PdfOptionalContentProperties? optionalContent = useDocumentWideObjects
            ? document.UncheckedOptionalContent
            : null;
        PdfDocumentOpenAction? openAction = useDocumentWideObjects
            ? document.UncheckedOpenAction
            : PdfPageRangeObjectFilter.FilterOpenActionByPageNumbers(document.UncheckedOpenAction, pageNumbers);

        cancellationToken.ThrowIfCancellationRequested();
        var pages = new List<PdfPageInfo>(pageNumbers.Length);
        var widgetsByPage = BuildFormWidgetsByPage(document.UncheckedFormFields, cancellationToken);
        for (int i = 0; i < pageNumbers.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            int pageNumber = pageNumbers[i];
            PdfReadPage page = document.Pages[pageNumber - 1];
            PdfPageGeometry geometry = page.GetGeometry();
            var (width, height) = PdfReadPage.GetPageSize(geometry);
            int rotation = page.GetRotationDegrees();
            var pageLinks = page.GetLinkAnnotationsUnchecked();
            var links = new List<PdfLinkAnnotation>(pageLinks.Count);
            for (int j = 0; j < pageLinks.Count; j++) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfLinkAnnotation link = pageLinks[j].WithPageNumber(pageNumber);
                if (link.DestinationPageObjectNumber.HasValue) {
                    link = link.WithDestinationPageNumber(document.GetPageNumberForObject(link.DestinationPageObjectNumber.Value));
                }

                links.Add(link);
            }

            var pageAnnotations = page.GetAnnotationsUnchecked();
            var annotations = new List<PdfAnnotation>(pageAnnotations.Count);
            for (int j = 0; j < pageAnnotations.Count; j++) {
                cancellationToken.ThrowIfCancellationRequested();
                annotations.Add(pageAnnotations[j].WithPageNumber(pageNumber));
            }

            var pageActions = page.GetPageActionsUnchecked();
            var actions = new List<PdfPageAction>(pageActions.Count);
            for (int j = 0; j < pageActions.Count; j++) {
                cancellationToken.ThrowIfCancellationRequested();
                actions.Add(pageActions[j].WithPageNumber(pageNumber));
            }

            widgetsByPage.TryGetValue(pageNumber, out IReadOnlyList<PdfFormWidget>? formWidgets);
            pages.Add(new PdfPageInfo(pageNumber, width, height, rotation, geometry, links, formWidgets, annotations, actions));
        }

        return new PdfDocumentInfo(pages.AsReadOnly(), document.UncheckedMetadata, outlines, pageLabels, namedDestinations, catalogActions, attachments, outputIntents, outputIntentsAreComplete, xmpMetadata, taggedContent, optionalContent, openAction, document.ViewerPreferences, formFields, document.UncheckedAcroFormDefaultAppearance, document.UncheckedAcroFormQuadding, document.UncheckedAcroFormXfa, document.UncheckedAcroFormNeedAppearances, document.UncheckedAcroFormSignatureFlags, document.Security, probe.HeaderVersion, document.CatalogPageMode, document.CatalogPageLayout, document.CatalogVersion, document.CatalogLanguage, document.Security.HasSignatures || probe.HasSignatures, probe.HasForms || document.UncheckedAcroFormXfa is not null, probe.HasAnnotations, probe.HasOutlines, probe.HasCatalogViewSettings, probe.HasPageLabels, probe.HasCatalogNameTrees, probe.HasNamedDestinations, probe.HasOpenActions, probe.HasViewerPreferences, probe.HasTaggedContent, probe.HasXmpMetadata, probe.HasCatalogUri, probe.HasOutputIntents, probe.HasEmbeddedFiles, probe.HasOptionalContent, probe.HasActiveContent, useDocumentWideObjects && document.HasOnlyWidgetOwnedActiveContent());
    }

    private static Dictionary<int, IReadOnlyList<PdfFormWidget>> BuildFormWidgetsByPage(
        IReadOnlyList<PdfFormField> fields,
        CancellationToken cancellationToken = default) {
        var grouped = new Dictionary<int, List<PdfFormWidget>>();
        for (int i = 0; i < fields.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<PdfFormWidget> widgets = fields[i].Widgets;
            for (int j = 0; j < widgets.Count; j++) {
                cancellationToken.ThrowIfCancellationRequested();
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
