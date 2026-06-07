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
        int viewerPreferencesId = 0,
        int structTreeRootId = 0,
        bool markInfo = false,
        string? openAction = null,
        string? pageMode = null,
        string? pageLayout = null,
        string? catalogUriBase = null) {
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

        if (structTreeRootId < 0) {
            throw new ArgumentOutOfRangeException(nameof(structTreeRootId), "PDF structure-tree root object number cannot be negative.");
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

        if (openAction != null && string.IsNullOrWhiteSpace(openAction)) {
            throw new ArgumentException("PDF catalog open action cannot be empty or whitespace.", nameof(openAction));
        }

        if (pageMode != null && string.IsNullOrWhiteSpace(pageMode)) {
            throw new ArgumentException("PDF catalog page mode cannot be empty or whitespace.", nameof(pageMode));
        }

        if (pageLayout != null && string.IsNullOrWhiteSpace(pageLayout)) {
            throw new ArgumentException("PDF catalog page layout cannot be empty or whitespace.", nameof(pageLayout));
        }

        if (catalogUriBase != null && string.IsNullOrWhiteSpace(catalogUriBase)) {
            throw new ArgumentException("PDF catalog URI base cannot be empty or whitespace.", nameof(catalogUriBase));
        }

        if (outlinesId > 0) {
            AppendReferenceEntry(sb, "Outlines", outlinesId);
            AppendNameEntry(sb, "PageMode", pageMode ?? "UseOutlines");
        } else if (pageMode != null) {
            AppendNameEntry(sb, "PageMode", pageMode);
        }

        if (pageLayout != null) {
            AppendNameEntry(sb, "PageLayout", pageLayout);
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

        if (catalogUriBase != null) {
            AppendCatalogUriEntry(sb, catalogUriBase);
        }

        if (openAction != null) {
            AppendOpenActionEntry(sb, openAction);
        }

        if (markInfo) {
            AppendMarkInfoEntry(sb);
        }

        if (structTreeRootId > 0) {
            AppendReferenceEntry(sb, "StructTreeRoot", structTreeRootId);
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

    internal static string BuildGeneratedOpenActionDestination(
        int pageObjectId,
        double destinationTop,
        PdfOpenActionDestinationMode destinationMode = PdfOpenActionDestinationMode.Xyz,
        double destinationLeft = 0d,
        double destinationBottom = 0d,
        double destinationRight = 0d) {
        if (pageObjectId < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageObjectId), "PDF open-action page object number must be positive.");
        }

        ValidateDestinationCoordinate(destinationTop, nameof(destinationTop));
        ValidateDestinationCoordinate(destinationLeft, nameof(destinationLeft));
        ValidateDestinationCoordinate(destinationBottom, nameof(destinationBottom));
        ValidateDestinationCoordinate(destinationRight, nameof(destinationRight));

        string left = destinationLeft.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
        string bottom = destinationBottom.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
        string right = destinationRight.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
        string top = destinationTop.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
        string pageReference = PdfSyntaxEscaper.IndirectReference(pageObjectId);

        switch (destinationMode) {
            case PdfOpenActionDestinationMode.Xyz:
                return "[" + pageReference + " /XYZ " + left + " " + top + " 0]";
            case PdfOpenActionDestinationMode.Fit:
                return "[" + pageReference + " /Fit]";
            case PdfOpenActionDestinationMode.FitHorizontal:
                return "[" + pageReference + " /FitH " + top + "]";
            case PdfOpenActionDestinationMode.FitVertical:
                return "[" + pageReference + " /FitV " + left + "]";
            case PdfOpenActionDestinationMode.FitRectangle:
                if (destinationRight <= destinationLeft) {
                    throw new ArgumentOutOfRangeException(nameof(destinationRight), "PDF open-action destination rectangle right coordinate must be greater than left coordinate.");
                }

                if (destinationTop <= destinationBottom) {
                    throw new ArgumentOutOfRangeException(nameof(destinationTop), "PDF open-action destination rectangle top coordinate must be greater than bottom coordinate.");
                }

                return "[" + pageReference + " /FitR " + left + " " + bottom + " " + right + " " + top + "]";
            case PdfOpenActionDestinationMode.FitBoundingBox:
                return "[" + pageReference + " /FitB]";
            case PdfOpenActionDestinationMode.FitBoundingBoxHorizontal:
                return "[" + pageReference + " /FitBH " + top + "]";
            case PdfOpenActionDestinationMode.FitBoundingBoxVertical:
                return "[" + pageReference + " /FitBV " + left + "]";
            default:
                throw new ArgumentOutOfRangeException(nameof(destinationMode), destinationMode, "PDF open-action destination mode is not supported.");
        }
    }

    private static void ValidateDestinationCoordinate(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "PDF open-action destination coordinate must be finite.");
        }
    }

    internal static string GetPageModeName(PdfCatalogPageMode pageMode) {
        Guard.CatalogPageMode(pageMode, nameof(pageMode));
        return pageMode switch {
            PdfCatalogPageMode.UseNone => "UseNone",
            PdfCatalogPageMode.UseOutlines => "UseOutlines",
            PdfCatalogPageMode.UseThumbs => "UseThumbs",
            PdfCatalogPageMode.FullScreen => "FullScreen",
            PdfCatalogPageMode.UseOC => "UseOC",
            PdfCatalogPageMode.UseAttachments => "UseAttachments",
            _ => throw new ArgumentOutOfRangeException(nameof(pageMode), "PDF catalog page mode is not supported.")
        };
    }

    internal static string GetPageLayoutName(PdfCatalogPageLayout pageLayout) {
        Guard.CatalogPageLayout(pageLayout, nameof(pageLayout));
        return pageLayout switch {
            PdfCatalogPageLayout.SinglePage => "SinglePage",
            PdfCatalogPageLayout.OneColumn => "OneColumn",
            PdfCatalogPageLayout.TwoColumnLeft => "TwoColumnLeft",
            PdfCatalogPageLayout.TwoColumnRight => "TwoColumnRight",
            PdfCatalogPageLayout.TwoPageLeft => "TwoPageLeft",
            PdfCatalogPageLayout.TwoPageRight => "TwoPageRight",
            _ => throw new ArgumentOutOfRangeException(nameof(pageLayout), "PDF catalog page layout is not supported.")
        };
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

    private static void AppendOpenActionEntry(StringBuilder sb, string openAction) {
        sb.Append(" /OpenAction ")
            .Append(openAction);
    }

    private static void AppendCatalogUriEntry(StringBuilder sb, string uriBase) {
        sb.Append(" /URI << /Base ")
            .Append(PdfSyntaxEscaper.LiteralString(uriBase))
            .Append(" >>");
    }

    private static void AppendMarkInfoEntry(StringBuilder sb) {
        sb.Append(" /MarkInfo << /Marked true >>");
    }
}
