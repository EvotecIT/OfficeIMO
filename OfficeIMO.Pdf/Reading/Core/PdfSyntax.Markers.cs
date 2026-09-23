namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    internal static void ThrowIfEncrypted(string trailerRaw) {
        if (ContainsPdfName(trailerRaw, "Encrypt")) {
            throw new NotSupportedException("Encrypted PDF files are not supported by OfficeIMO.Pdf yet.");
        }
    }

    internal static bool HasEncryptionMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ReadDocumentSecurityInfo(pdf).HasEncryption;
    }

    internal static bool HasSignatureMarkers(byte[] pdf, PdfLoadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, options, "ByteRange", "SigFlags", "Sig");
    }

    internal static bool HasFormMarkers(byte[] pdf) {
        return HasFormMarkers(pdf, null);
    }

    private static bool HasFormMarkers(byte[] pdf, PdfLoadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));

        if (options is not null) {
            return ContainsParsedOrFallbackPdfName(pdf, options, "AcroForm", "Fields", "FT", "XFA");
        }

        return ContainsParsedOrFallbackPdfName(pdf, "AcroForm", "Fields", "FT", "XFA");
    }

    internal static bool HasAnnotationMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "Annots", "Annot");
    }

    internal static bool HasOutlineMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "Outlines", "UseOutlines");
    }

    internal static bool HasUnsupportedOutlineRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasOutlineMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("Outlines", out var outlines)) {
                return false;
            }

            return !IsSupportedOutlineGraph(objects, outlines, new HashSet<int>());
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasCatalogViewSettingMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "PageMode", "PageLayout");
    }

    internal static bool HasPageLabelMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "PageLabels");
    }

    internal static bool HasUnsupportedPageLabelRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasPageLabelMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("PageLabels", out var pageLabels)) {
                return catalog is null;
            }

            return !IsSupportedPageLabelTree(objects, pageLabels);
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasNamedDestinationMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "Dests");
    }

    internal static bool HasCatalogNameTreeMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "Names");
    }

    internal static bool HasUnsupportedCatalogNameTreeRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasCatalogNameTreeMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("Names", out var names)) {
                return false;
            }

            PdfDictionary? namesDictionary = ResolveObject(objects, names) as PdfDictionary;
            if (namesDictionary is null) {
                return true;
            }

            foreach (var key in namesDictionary.Items.Keys) {
                if (string.Equals(key, "Dests", StringComparison.Ordinal) ||
                    string.Equals(key, "EmbeddedFiles", StringComparison.Ordinal) ||
                    string.Equals(key, "JavaScript", StringComparison.Ordinal)) {
                    continue;
                }

                return true;
            }

            return false;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasUnsupportedNamedDestinationRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasNamedDestinationMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfReadLimits limits = options?.Limits ?? PdfLoadOptions.Default.Limits;
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null) {
                return true;
            }

            if (catalog.Items.ContainsKey("Dests")) {
                return false;
            }

            if (catalog.Items.TryGetValue("Names", out var names)) {
                PdfDictionary? namesDictionary = ResolveObject(objects, names) as PdfDictionary;
                if (namesDictionary is null) {
                    return true;
                }

                if (namesDictionary.Items.ContainsKey("Dests")) {
                    return !TryGetNamedDestinationNameTree(objects, names, out var namedDestinationTree) ||
                        !IsSupportedNamedDestinationNameTree(objects, namedDestinationTree, limits);
                }
            }

            return false;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasOpenActionMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "OpenAction");
    }

    internal static bool HasUnsupportedOpenActionRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasOpenActionMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("OpenAction", out var openAction)) {
                return catalog is null;
            }

            PdfObject? resolved = ResolveObject(objects, openAction);
            if (resolved is PdfArray array &&
                IsDestinationForKnownPage(objects, array)) {
                return false;
            }

            if (resolved is PdfDictionary dictionary &&
                IsSupportedGoToActionDictionary(objects, dictionary)) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasViewerPreferenceMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "ViewerPreferences");
    }

    internal static bool HasUnsupportedViewerPreferenceRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasViewerPreferenceMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("ViewerPreferences", out var viewerPreferences)) {
                return catalog is null;
            }

            PdfObject? resolved = ResolveObject(objects, viewerPreferences);
            if (resolved is PdfDictionary dictionary &&
                IsSimpleCatalogDictionary(dictionary)) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasTaggedContentMarkers(byte[] pdf) {
        return HasTaggedContentMarkers(pdf, null);
    }

    private static bool HasTaggedContentMarkers(byte[] pdf, PdfLoadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));

        if (options is not null) {
            return ContainsParsedOrFallbackPdfName(pdf, options, "MarkInfo", "StructTreeRoot", "ParentTree", "StructElem");
        }

        return ContainsParsedOrFallbackPdfName(pdf, "MarkInfo", "StructTreeRoot", "ParentTree", "StructElem");
    }

    internal static bool HasXmpMetadataMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "Metadata");
    }

    internal static bool HasUnsupportedXmpMetadataRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasXmpMetadataMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("Metadata", out var xmpMetadata)) {
                return catalog is null;
            }

            if (IsSupportedCatalogXmpMetadataStream(objects, xmpMetadata)) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasCatalogUriMarkers(byte[] pdf, PdfLoadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        if (!ContainsPdfName(text, "URI")) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            return catalog?.Items.ContainsKey("URI") == true;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasUnsupportedCatalogUriRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasCatalogUriMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("URI", out var catalogUri)) {
                return catalog is null;
            }

            PdfObject? resolved = ResolveObject(objects, catalogUri);
            if (resolved is PdfDictionary dictionary &&
                IsSimpleCatalogDictionary(dictionary)) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasOutputIntentMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "OutputIntents", "OutputIntent");
    }

    internal static bool HasUnsupportedOutputIntentRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasOutputIntentMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("OutputIntents", out var outputIntents)) {
                return catalog is null;
            }

            if (IsSupportedCatalogMetadataGraph(objects, outputIntents, new HashSet<int>())) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasEmbeddedFileMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "EmbeddedFiles", "Filespec", "EmbeddedFile", "AF");
    }

    internal static bool HasUnsupportedEmbeddedFileRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasEmbeddedFileMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null) {
                return true;
            }

            if (catalog.Items.TryGetValue("Names", out var names)) {
                PdfDictionary? namesDictionary = ResolveObject(objects, names) as PdfDictionary;
                if (namesDictionary is null) {
                    return true;
                }

                if (namesDictionary.Items.ContainsKey("EmbeddedFiles")) {
                    if (!TryGetEmbeddedFilesNameTree(objects, names, out var embeddedFiles) ||
                        !IsSupportedCatalogMetadataGraph(objects, embeddedFiles, new HashSet<int>())) {
                        return true;
                    }
                }
            }

            if (catalog.Items.TryGetValue("AF", out var associatedFiles)) {
                if (!IsSupportedCatalogMetadataGraph(objects, associatedFiles, new HashSet<int>())) {
                    return true;
                }
            }

            return false;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasOptionalContentMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ContainsParsedOrFallbackPdfName(pdf, "OCProperties", "OCGs", "OCG", "OCMD");
    }

    internal static bool HasUnsupportedOptionalContentRewriteMarkers(byte[] pdf, PdfLoadOptions? options = null, PdfRewriteMarkerSource? source = null) {
        if (source is null && options is null && !HasOptionalContentMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseRewriteMarkerObjects(pdf, options, source);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("OCProperties", out var optionalContent)) {
                return catalog is null;
            }

            if (IsSupportedCatalogMetadataGraph(objects, optionalContent, new HashSet<int>())) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    private static (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) ParseRewriteMarkerObjects(
        byte[] pdf,
        PdfLoadOptions? options,
        PdfRewriteMarkerSource? source) => source?.Parse() ?? ParseObjects(pdf, options);

    internal static bool HasActiveContentMarkers(byte[] pdf) {
        return HasActiveContentMarkers(pdf, null);
    }

    private static bool HasActiveContentMarkers(byte[] pdf, PdfLoadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));

        if (options is not null) {
            return ContainsParsedOrFallbackPdfName(pdf, options, PdfActiveContentPolicy.MarkerNames);
        }

        return ContainsParsedOrFallbackPdfName(pdf, PdfActiveContentPolicy.MarkerNames);
    }

    internal static string? GetHeaderVersion(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        const int maximumVersionTokenLength = 16;

        if (pdf.Length < 8 ||
            pdf[0] != (byte)'%' ||
            pdf[1] != (byte)'P' ||
            pdf[2] != (byte)'D' ||
            pdf[3] != (byte)'F' ||
            pdf[4] != (byte)'-') {
            return null;
        }

        int start = 5;
        int end = start;
        while (end < pdf.Length && end - start <= maximumVersionTokenLength) {
            byte value = pdf[end];
            if (value == (byte)'\r' || value == (byte)'\n' || value == (byte)' ' || value == (byte)'\t') {
                break;
            }

            end++;
        }

        // A PDF version is a short token. Do not scan or allocate an entire malformed input line.
        if (end - start > maximumVersionTokenLength) return null;
        return end > start ? PdfEncoding.Latin1GetString(pdf, start, end - start) : null;
    }

}
