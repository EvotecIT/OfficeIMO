namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    internal static void ThrowIfEncrypted(string trailerRaw) {
        if (ContainsPdfName(trailerRaw, "Encrypt")) {
            throw new NotSupportedException("Encrypted PDF files are not supported by OfficeIMO.Pdf yet.");
        }
    }

    internal static void ThrowIfUnsafeForRewrite(byte[] pdf) {
        ThrowIfUnsafeForRewrite(pdf, allowEncryption: false);
    }

    internal static void ThrowIfUnsafeForRewrite(byte[] pdf, PdfReadOptions? options) {
        ThrowIfUnsafeForRewrite(pdf, allowEncryption: options?.Password is not null || CanOpenEncryptedPdfWithEmptyPassword(pdf), options);
    }

    internal static void ThrowIfUnsafeForRewrite(byte[] pdf, bool allowEncryption) {
        ThrowIfUnsafeForRewrite(pdf, allowEncryption, options: null);
    }

    private static void ThrowIfUnsafeForRewrite(byte[] pdf, bool allowEncryption, PdfReadOptions? options) {
        if (!allowEncryption && HasEncryptionMarkers(pdf)) {
            throw new NotSupportedException("Encrypted PDF files are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasSignatureMarkers(pdf)) {
            throw new NotSupportedException("Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasFormMarkers(pdf, options)) {
            throw new NotSupportedException("PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedOutlineRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedPageLabelRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF page labels are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedCatalogNameTreeRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedNamedDestinationRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedOpenActionRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedViewerPreferenceRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasTaggedContentMarkers(pdf, options)) {
            throw new NotSupportedException("PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedXmpMetadataRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedCatalogUriRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedOutputIntentRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedEmbeddedFileRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedOptionalContentRewriteMarkers(pdf, options)) {
            throw new NotSupportedException("PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasActiveContentMarkers(pdf, options)) {
            throw new NotSupportedException("PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
        }
    }

    private static bool CanOpenEncryptedPdfWithEmptyPassword(byte[] pdf) {
        if (!HasEncryptionMarkers(pdf)) {
            return false;
        }

        try {
            PdfReadDocument.Open(pdf, new PdfReadOptions { Password = string.Empty });
            return true;
        } catch (PdfEncryptionException) {
            return false;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return false;
        }
    }

    internal static bool HasEncryptionMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        return ReadDocumentSecurityInfo(pdf).HasEncryption;
    }

    internal static bool HasSignatureMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "ByteRange", "SigFlags", "Sig") ||
            ContainsAnyParsedPdfName(pdf, "ByteRange", "SigFlags", "Sig");
    }

    internal static bool HasFormMarkers(byte[] pdf) {
        return HasFormMarkers(pdf, null);
    }

    private static bool HasFormMarkers(byte[] pdf, PdfReadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));

        if (options is not null) {
            return ContainsAnyParsedPdfName(pdf, options, "AcroForm", "Fields", "FT", "XFA");
        }

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "AcroForm", "Fields", "FT", "XFA") ||
            ContainsAnyParsedPdfName(pdf, "AcroForm", "Fields", "FT", "XFA");
    }

    internal static bool HasAnnotationMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "Annots", "Annot") ||
            ContainsAnyParsedPdfName(pdf, "Annots", "Annot");
    }

    internal static bool HasOutlineMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "Outlines", "UseOutlines") ||
            ContainsAnyParsedPdfName(pdf, "Outlines", "UseOutlines");
    }

    internal static bool HasUnsupportedOutlineRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasOutlineMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "PageMode", "PageLayout") ||
            ContainsAnyParsedPdfName(pdf, "PageMode", "PageLayout");
    }

    internal static bool HasPageLabelMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "PageLabels") ||
            ContainsAnyParsedPdfName(pdf, "PageLabels");
    }

    internal static bool HasUnsupportedPageLabelRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasPageLabelMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "Dests") ||
            ContainsAnyParsedPdfName(pdf, "Dests");
    }

    internal static bool HasCatalogNameTreeMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "Names") ||
            ContainsAnyParsedPdfName(pdf, "Names");
    }

    internal static bool HasUnsupportedCatalogNameTreeRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasCatalogNameTreeMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

    internal static bool HasUnsupportedNamedDestinationRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasNamedDestinationMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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
                        !IsSupportedNamedDestinationNameTree(objects, namedDestinationTree);
                }
            }

            return false;
        } catch (Exception ex) when (ex is not PdfEncryptionException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasOpenActionMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "OpenAction") ||
            ContainsAnyParsedPdfName(pdf, "OpenAction");
    }

    internal static bool HasUnsupportedOpenActionRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasOpenActionMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "ViewerPreferences") ||
            ContainsAnyParsedPdfName(pdf, "ViewerPreferences");
    }

    internal static bool HasUnsupportedViewerPreferenceRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasViewerPreferenceMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

    private static bool HasTaggedContentMarkers(byte[] pdf, PdfReadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));

        if (options is not null) {
            return ContainsAnyParsedPdfName(pdf, options, "MarkInfo", "StructTreeRoot", "ParentTree", "StructElem");
        }

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "MarkInfo", "StructTreeRoot", "ParentTree", "StructElem") ||
            ContainsAnyParsedPdfName(pdf, "MarkInfo", "StructTreeRoot", "ParentTree", "StructElem");
    }

    internal static bool HasXmpMetadataMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "Metadata") ||
            ContainsAnyParsedPdfName(pdf, "Metadata");
    }

    internal static bool HasUnsupportedXmpMetadataRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasXmpMetadataMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

    internal static bool HasCatalogUriMarkers(byte[] pdf, PdfReadOptions? options = null) {
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

    internal static bool HasUnsupportedCatalogUriRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasCatalogUriMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "OutputIntents") ||
            ContainsPdfName(text, "OutputIntent");
    }

    internal static bool HasUnsupportedOutputIntentRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasOutputIntentMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "EmbeddedFiles", "Filespec", "EmbeddedFile", "AF") ||
            ContainsAnyParsedPdfName(pdf, "EmbeddedFiles", "Filespec", "EmbeddedFile", "AF");
    }

    internal static bool HasUnsupportedEmbeddedFileRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasEmbeddedFileMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "OCProperties", "OCGs", "OCG", "OCMD") ||
            ContainsAnyParsedPdfName(pdf, "OCProperties", "OCGs", "OCG", "OCMD");
    }

    internal static bool HasUnsupportedOptionalContentRewriteMarkers(byte[] pdf, PdfReadOptions? options = null) {
        if (options is null && !HasOptionalContentMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
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

    internal static bool HasActiveContentMarkers(byte[] pdf) {
        return HasActiveContentMarkers(pdf, null);
    }

    private static bool HasActiveContentMarkers(byte[] pdf, PdfReadOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));

        if (options is not null) {
            return ContainsAnyParsedPdfName(pdf, options, "JavaScript", "JS", "AA", "Launch", "SubmitForm", "RichMedia");
        }

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "JavaScript", "JS", "AA", "Launch", "SubmitForm", "RichMedia") ||
            ContainsAnyParsedPdfName(pdf, "JavaScript", "JS", "AA", "Launch", "SubmitForm", "RichMedia");
    }

    internal static string? GetHeaderVersion(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

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
        while (end < pdf.Length) {
            byte value = pdf[end];
            if (value == (byte)'\r' || value == (byte)'\n' || value == (byte)' ' || value == (byte)'\t') {
                break;
            }

            end++;
        }

        return end > start ? PdfEncoding.Latin1GetString(pdf, start, end - start) : null;
    }

}
