namespace OfficeIMO.Pdf;

internal static class PdfViewerPreferenceDictionaryBuilder {
    internal static string BuildGeneratedViewerPreferencesDictionary(PdfViewerPreferencesOptions preferences) {
        Guard.NotNull(preferences, nameof(preferences));
        if (!preferences.HasAny) {
            throw new ArgumentException("At least one PDF viewer preference must be configured.", nameof(preferences));
        }

        var sb = new StringBuilder();
        sb.Append("<<");
        AppendBooleanEntry(sb, "HideToolbar", preferences.HideToolbar);
        AppendBooleanEntry(sb, "HideMenubar", preferences.HideMenubar);
        AppendBooleanEntry(sb, "HideWindowUI", preferences.HideWindowUI);
        AppendBooleanEntry(sb, "FitWindow", preferences.FitWindow);
        AppendBooleanEntry(sb, "CenterWindow", preferences.CenterWindow);
        AppendBooleanEntry(sb, "DisplayDocTitle", preferences.DisplayDocTitle);
        AppendBooleanEntry(sb, "PickTrayByPDFSize", preferences.PickTrayByPdfSize);
        if (preferences.NonFullScreenPageMode.HasValue) {
            AppendNameEntry(sb, "NonFullScreenPageMode", GetNonFullScreenPageModeName(preferences.NonFullScreenPageMode.Value));
        }

        if (preferences.Direction.HasValue) {
            AppendNameEntry(sb, "Direction", GetDirectionName(preferences.Direction.Value));
        }

        if (preferences.PrintScaling.HasValue) {
            AppendNameEntry(sb, "PrintScaling", GetPrintScalingName(preferences.PrintScaling.Value));
        }

        if (preferences.Duplex.HasValue) {
            AppendNameEntry(sb, "Duplex", GetDuplexName(preferences.Duplex.Value));
        }

        if (preferences.ViewArea.HasValue) {
            AppendNameEntry(sb, "ViewArea", GetPageBoundaryBoxName(preferences.ViewArea.Value));
        }

        if (preferences.ViewClip.HasValue) {
            AppendNameEntry(sb, "ViewClip", GetPageBoundaryBoxName(preferences.ViewClip.Value));
        }

        if (preferences.PrintArea.HasValue) {
            AppendNameEntry(sb, "PrintArea", GetPageBoundaryBoxName(preferences.PrintArea.Value));
        }

        if (preferences.PrintClip.HasValue) {
            AppendNameEntry(sb, "PrintClip", GetPageBoundaryBoxName(preferences.PrintClip.Value));
        }

        if (preferences.NumCopies.HasValue) {
            AppendPositiveIntegerEntry(sb, "NumCopies", preferences.NumCopies.Value);
        }

        if (preferences.PrintPageRanges.Count > 0) {
            AppendPrintPageRangeEntry(sb, preferences.PrintPageRanges);
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    internal static string GetNonFullScreenPageModeName(PdfNonFullScreenPageMode pageMode) {
        Guard.NonFullScreenPageMode(pageMode, nameof(pageMode));
        return pageMode switch {
            PdfNonFullScreenPageMode.UseNone => "UseNone",
            PdfNonFullScreenPageMode.UseOutlines => "UseOutlines",
            PdfNonFullScreenPageMode.UseThumbs => "UseThumbs",
            PdfNonFullScreenPageMode.UseOC => "UseOC",
            _ => throw new ArgumentOutOfRangeException(nameof(pageMode), "PDF non-full-screen page mode is not supported.")
        };
    }

    internal static string GetDirectionName(PdfViewerDirection direction) {
        Guard.ViewerDirection(direction, nameof(direction));
        return direction switch {
            PdfViewerDirection.LeftToRight => "L2R",
            PdfViewerDirection.RightToLeft => "R2L",
            _ => throw new ArgumentOutOfRangeException(nameof(direction), "PDF viewer direction is not supported.")
        };
    }

    internal static string GetPrintScalingName(PdfPrintScaling printScaling) {
        Guard.PrintScaling(printScaling, nameof(printScaling));
        return printScaling switch {
            PdfPrintScaling.AppDefault => "AppDefault",
            PdfPrintScaling.None => "None",
            _ => throw new ArgumentOutOfRangeException(nameof(printScaling), "PDF print scaling is not supported.")
        };
    }

    internal static string GetDuplexName(PdfDuplexMode duplex) {
        Guard.DuplexMode(duplex, nameof(duplex));
        return duplex switch {
            PdfDuplexMode.Simplex => "Simplex",
            PdfDuplexMode.DuplexFlipShortEdge => "DuplexFlipShortEdge",
            PdfDuplexMode.DuplexFlipLongEdge => "DuplexFlipLongEdge",
            _ => throw new ArgumentOutOfRangeException(nameof(duplex), "PDF duplex mode is not supported.")
        };
    }

    internal static string GetPageBoundaryBoxName(PdfPageBoundaryBox boundaryBox) {
        Guard.PageBoundaryBox(boundaryBox, nameof(boundaryBox));
        return boundaryBox switch {
            PdfPageBoundaryBox.MediaBox => "MediaBox",
            PdfPageBoundaryBox.CropBox => "CropBox",
            PdfPageBoundaryBox.BleedBox => "BleedBox",
            PdfPageBoundaryBox.TrimBox => "TrimBox",
            PdfPageBoundaryBox.ArtBox => "ArtBox",
            _ => throw new ArgumentOutOfRangeException(nameof(boundaryBox), "PDF page boundary box is not supported.")
        };
    }

    private static void AppendBooleanEntry(StringBuilder sb, string key, bool? value) {
        if (!value.HasValue) {
            return;
        }

        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(value.Value ? " true" : " false");
    }

    private static void AppendNameEntry(StringBuilder sb, string key, string value) {
        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(" /")
            .Append(PdfSyntaxEscaper.Name(value));
    }

    private static void AppendPositiveIntegerEntry(StringBuilder sb, string key, int value) {
        Guard.PositiveInteger(value, key);
        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(' ')
            .Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    private static void AppendPrintPageRangeEntry(StringBuilder sb, IReadOnlyList<PdfPrintPageRange> ranges) {
        sb.Append(" /PrintPageRange [");
        for (int i = 0; i < ranges.Count; i++) {
            if (i > 0) {
                sb.Append(' ');
            }

            PdfPrintPageRange range = ranges[i];
            sb.Append((range.StartPageNumber - 1).ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(' ')
                .Append((range.EndPageNumber - 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        sb.Append(']');
    }
}
