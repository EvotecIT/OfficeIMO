namespace OfficeIMO.Pdf;

internal static class PdfCatalogDictionaryBuilder {
    internal static string BuildGeneratedCatalogDictionary(
        int pagesId,
        int outlinesId,
        int namedDestinationsId = 0,
        int acroFormId = 0,
        int metadataId = 0,
        int outputIntentId = 0,
        string? language = null,
        int embeddedFilesNameTreeId = 0,
        IReadOnlyList<int>? associatedFileIds = null,
        int pageLabelsId = 0,
        int viewerPreferencesId = 0) {
        var sb = new StringBuilder();
        AppendCatalogStart(sb, pagesId);

        if (outlinesId < 0) {
            throw new ArgumentOutOfRangeException(nameof(outlinesId), "PDF outline object number cannot be negative.");
        }

        if (namedDestinationsId < 0) {
            throw new ArgumentOutOfRangeException(nameof(namedDestinationsId), "PDF named destinations object number cannot be negative.");
        }

        if (acroFormId < 0) {
            throw new ArgumentOutOfRangeException(nameof(acroFormId), "PDF AcroForm object number cannot be negative.");
        }

        if (metadataId < 0) {
            throw new ArgumentOutOfRangeException(nameof(metadataId), "PDF metadata object number cannot be negative.");
        }

        if (outputIntentId < 0) {
            throw new ArgumentOutOfRangeException(nameof(outputIntentId), "PDF output intent object number cannot be negative.");
        }

        if (embeddedFilesNameTreeId < 0) {
            throw new ArgumentOutOfRangeException(nameof(embeddedFilesNameTreeId), "PDF embedded-files name-tree object number cannot be negative.");
        }

        if (pageLabelsId < 0) {
            throw new ArgumentOutOfRangeException(nameof(pageLabelsId), "PDF page-label object number cannot be negative.");
        }

        if (viewerPreferencesId < 0) {
            throw new ArgumentOutOfRangeException(nameof(viewerPreferencesId), "PDF viewer-preferences object number cannot be negative.");
        }

        if (associatedFileIds != null) {
            for (int i = 0; i < associatedFileIds.Count; i++) {
                if (associatedFileIds[i] < 1) {
                    throw new ArgumentOutOfRangeException(nameof(associatedFileIds), "PDF associated-file object numbers must be positive.");
                }
            }
        }

        if (language != null && string.IsNullOrWhiteSpace(language)) {
            throw new ArgumentException("PDF catalog language cannot be empty or whitespace.", nameof(language));
        }

        if (outlinesId > 0) {
            AppendReferenceEntry(sb, "Outlines", outlinesId);
            AppendNameEntry(sb, "PageMode", "UseOutlines");
        }

        if (language != null) {
            AppendTextStringEntry(sb, "Lang", language);
        }

        if (pageLabelsId > 0) {
            AppendReferenceEntry(sb, "PageLabels", pageLabelsId);
        }

        if (namedDestinationsId > 0 || embeddedFilesNameTreeId > 0) {
            AppendNamesEntry(sb, namedDestinationsId, embeddedFilesNameTreeId);
        }

        if (acroFormId > 0) {
            AppendReferenceEntry(sb, "AcroForm", acroFormId);
        }

        if (viewerPreferencesId > 0) {
            AppendReferenceEntry(sb, "ViewerPreferences", viewerPreferencesId);
        }

        if (metadataId > 0) {
            AppendReferenceEntry(sb, "Metadata", metadataId);
        }

        if (outputIntentId > 0) {
            AppendOutputIntentEntry(sb, outputIntentId);
        }

        if (associatedFileIds != null && associatedFileIds.Count > 0) {
            AppendAssociatedFilesEntry(sb, associatedFileIds);
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    internal static void AppendCatalogStart(StringBuilder sb, int pagesId) {
        Guard.NotNull(sb, nameof(sb));
        sb.Append("<< /Type /Catalog /Pages ")
            .Append(PdfSyntaxEscaper.IndirectReference(pagesId));
    }

    internal static void AppendNameEntry(StringBuilder sb, string key, string value) {
        Guard.NotNull(sb, nameof(sb));
        Guard.NotNullOrWhiteSpace(key, nameof(key));
        Guard.NotNullOrWhiteSpace(value, nameof(value));
        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(" /")
            .Append(PdfSyntaxEscaper.Name(value));
    }

    internal static void AppendReferenceEntry(StringBuilder sb, string key, int objectNumber, int generation = 0) {
        Guard.NotNull(sb, nameof(sb));
        Guard.NotNullOrWhiteSpace(key, nameof(key));
        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(' ')
            .Append(PdfSyntaxEscaper.IndirectReference(objectNumber, generation));
    }

    internal static void AppendTextStringEntry(StringBuilder sb, string key, string value) {
        Guard.NotNull(sb, nameof(sb));
        Guard.NotNullOrWhiteSpace(key, nameof(key));
        Guard.NotNullOrWhiteSpace(value, nameof(value));
        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(' ')
            .Append(PdfSyntaxEscaper.TextString(value));
    }

    private static void AppendNamesEntry(StringBuilder sb, int namedDestinationsId, int embeddedFilesNameTreeId) {
        sb.Append(" /Names <<");
        if (namedDestinationsId > 0) {
            sb.Append(" /Dests ")
                .Append(PdfSyntaxEscaper.IndirectReference(namedDestinationsId));
        }

        if (embeddedFilesNameTreeId > 0) {
            sb.Append(" /EmbeddedFiles ")
                .Append(PdfSyntaxEscaper.IndirectReference(embeddedFilesNameTreeId));
        }

        sb.Append(" >>");
    }

    private static void AppendOutputIntentEntry(StringBuilder sb, int objectNumber) {
        sb.Append(" /OutputIntents [")
            .Append(PdfSyntaxEscaper.IndirectReference(objectNumber))
            .Append(']');
    }

    private static void AppendAssociatedFilesEntry(StringBuilder sb, IReadOnlyList<int> objectNumbers) {
        sb.Append(" /AF [");
        for (int i = 0; i < objectNumbers.Count; i++) {
            if (i > 0) {
                sb.Append(' ');
            }

            sb.Append(PdfSyntaxEscaper.IndirectReference(objectNumbers[i]));
        }

        sb.Append(']');
    }
}
