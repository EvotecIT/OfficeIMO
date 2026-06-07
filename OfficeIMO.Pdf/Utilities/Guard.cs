using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

internal static class Guard {
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void NotNull<T>(T? value, string paramName) where T : class {
        if (value is null) throw new System.ArgumentNullException(paramName, $"Parameter '{paramName}' cannot be null.");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void NotNullOrWhiteSpace(string? value, string paramName) {
        if (value is null)
            throw new System.ArgumentNullException(paramName, $"Parameter '{paramName}' cannot be null.");

        if (string.IsNullOrWhiteSpace(value))
            throw new System.ArgumentException($"Parameter '{paramName}' cannot be empty or whitespace.", paramName);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void AbsoluteUri(string? value, string paramName) {
        NotNullOrWhiteSpace(value, paramName);
        if (!System.Uri.TryCreate(value, System.UriKind.Absolute, out _))
            throw new System.ArgumentException($"Parameter '{paramName}' must be an absolute URI.", paramName);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void OptionalAbsoluteUri(string? value, string paramName) {
        if (value is null)
            return;

        AbsoluteUri(value, paramName);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void UriAction(string? value, string paramName) {
        NotNullOrWhiteSpace(value, paramName);
        if (!IsUriAction(value)) {
            throw new System.ArgumentException($"Parameter '{paramName}' must be an absolute URI or a relative URI action target.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void OptionalUriAction(string? value, string paramName) {
        if (value is null)
            return;

        UriAction(value, paramName);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsUriAction(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string uriAction = value!;
        for (int i = 0; i < uriAction.Length; i++) {
            if (char.IsControl(uriAction[i])) {
                return false;
            }
        }

        return System.Uri.TryCreate(uriAction, System.UriKind.RelativeOrAbsolute, out _);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void NotNullOrEmpty(byte[]? value, string paramName) {
        if (value is null)
            throw new System.ArgumentNullException(paramName, $"Parameter '{paramName}' cannot be null.");

        if (value.Length == 0)
            throw new System.ArgumentException($"Parameter '{paramName}' cannot be empty.", paramName);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void Positive(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0)
            throw new System.ArgumentOutOfRangeException(paramName, value, $"Parameter '{paramName}' must be a finite positive number.");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void NonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0)
            throw new System.ArgumentOutOfRangeException(paramName, value, $"Parameter '{paramName}' must be a finite non-negative number.");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void PositiveInteger(int value, string paramName) {
        if (value < 1) {
            throw new System.ArgumentOutOfRangeException(paramName, value, $"Parameter '{paramName}' must be a positive integer.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void LeftCenterRightAlign(PdfAlign value, string paramName, string context) {
        if (value != PdfAlign.Left && value != PdfAlign.Center && value != PdfAlign.Right) {
            throw new System.ArgumentException($"{context} alignment must be Left, Center, or Right.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void ParagraphAlign(PdfAlign value, string paramName, string context) {
        if (value != PdfAlign.Left && value != PdfAlign.Center && value != PdfAlign.Right && value != PdfAlign.Justify) {
            throw new System.ArgumentException($"{context} alignment must be Left, Center, Right, or Justify.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void TableColumnAlign(PdfColumnAlign value, string paramName) {
        if (value != PdfColumnAlign.Left && value != PdfColumnAlign.Center && value != PdfColumnAlign.Right) {
            throw new System.ArgumentException("Table column alignments must be Left, Center, or Right.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void TableCellVerticalAlign(PdfCellVerticalAlign value, string paramName) {
        if (value != PdfCellVerticalAlign.Top && value != PdfCellVerticalAlign.Middle && value != PdfCellVerticalAlign.Bottom) {
            throw new System.ArgumentException("Table vertical alignments must be defined PDF cell vertical alignment values.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void PageOrientation(PdfPageOrientation value, string paramName) {
        if (value != PdfPageOrientation.Portrait && value != PdfPageOrientation.Landscape) {
            throw new System.ArgumentException("PDF page orientation must be Portrait or Landscape.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void PageNumberStyle(PdfPageNumberStyle value, string paramName) {
        if (value != PdfPageNumberStyle.Arabic &&
            value != PdfPageNumberStyle.LowerRoman &&
            value != PdfPageNumberStyle.UpperRoman &&
            value != PdfPageNumberStyle.LowerLetter &&
            value != PdfPageNumberStyle.UpperLetter) {
            throw new System.ArgumentException("PDF page number style must be Arabic, LowerRoman, UpperRoman, LowerLetter, or UpperLetter.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void CatalogPageMode(PdfCatalogPageMode value, string paramName) {
        if (value != PdfCatalogPageMode.UseNone &&
            value != PdfCatalogPageMode.UseOutlines &&
            value != PdfCatalogPageMode.UseThumbs &&
            value != PdfCatalogPageMode.FullScreen &&
            value != PdfCatalogPageMode.UseOC &&
            value != PdfCatalogPageMode.UseAttachments) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF catalog page mode must be UseNone, UseOutlines, UseThumbs, FullScreen, UseOC, or UseAttachments.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void CatalogPageLayout(PdfCatalogPageLayout value, string paramName) {
        if (value != PdfCatalogPageLayout.SinglePage &&
            value != PdfCatalogPageLayout.OneColumn &&
            value != PdfCatalogPageLayout.TwoColumnLeft &&
            value != PdfCatalogPageLayout.TwoColumnRight &&
            value != PdfCatalogPageLayout.TwoPageLeft &&
            value != PdfCatalogPageLayout.TwoPageRight) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF catalog page layout must be SinglePage, OneColumn, TwoColumnLeft, TwoColumnRight, TwoPageLeft, or TwoPageRight.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void NonFullScreenPageMode(PdfNonFullScreenPageMode value, string paramName) {
        if (value != PdfNonFullScreenPageMode.UseNone &&
            value != PdfNonFullScreenPageMode.UseOutlines &&
            value != PdfNonFullScreenPageMode.UseThumbs &&
            value != PdfNonFullScreenPageMode.UseOC) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF non-full-screen page mode must be UseNone, UseOutlines, UseThumbs, or UseOC.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void ViewerDirection(PdfViewerDirection value, string paramName) {
        if (value != PdfViewerDirection.LeftToRight &&
            value != PdfViewerDirection.RightToLeft) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF viewer direction must be LeftToRight or RightToLeft.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void PrintScaling(PdfPrintScaling value, string paramName) {
        if (value != PdfPrintScaling.AppDefault &&
            value != PdfPrintScaling.None) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF print scaling must be AppDefault or None.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void DuplexMode(PdfDuplexMode value, string paramName) {
        if (value != PdfDuplexMode.Simplex &&
            value != PdfDuplexMode.DuplexFlipShortEdge &&
            value != PdfDuplexMode.DuplexFlipLongEdge) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF duplex mode must be Simplex, DuplexFlipShortEdge, or DuplexFlipLongEdge.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void FormFieldTextAlignment(PdfFormFieldTextAlignment value, string paramName) {
        if (value != PdfFormFieldTextAlignment.Left &&
            value != PdfFormFieldTextAlignment.Center &&
            value != PdfFormFieldTextAlignment.Right) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF form field text alignment must be Left, Center, or Right.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void PageBoundaryBox(PdfPageBoundaryBox value, string paramName) {
        if (value != PdfPageBoundaryBox.MediaBox &&
            value != PdfPageBoundaryBox.CropBox &&
            value != PdfPageBoundaryBox.BleedBox &&
            value != PdfPageBoundaryBox.TrimBox &&
            value != PdfPageBoundaryBox.ArtBox) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF page boundary box must be MediaBox, CropBox, BleedBox, TrimBox, or ArtBox.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void ComplianceProfile(PdfComplianceProfile value, string paramName) {
        if (value != PdfComplianceProfile.None &&
            value != PdfComplianceProfile.PdfA2B &&
            value != PdfComplianceProfile.PdfA2U &&
            value != PdfComplianceProfile.PdfA2A &&
            value != PdfComplianceProfile.PdfA3B &&
            value != PdfComplianceProfile.PdfA3U &&
            value != PdfComplianceProfile.PdfA3A &&
            value != PdfComplianceProfile.PdfUa1 &&
            value != PdfComplianceProfile.FacturX &&
            value != PdfComplianceProfile.Zugferd) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF compliance profile must be None, PdfA2B, PdfA2U, PdfA2A, PdfA3B, PdfA3U, PdfA3A, PdfUa1, FacturX, or Zugferd.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void FileVersion(PdfFileVersion value, string paramName) {
        if (value != PdfFileVersion.Pdf14 &&
            value != PdfFileVersion.Pdf15 &&
            value != PdfFileVersion.Pdf16 &&
            value != PdfFileVersion.Pdf17) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF file version must be Pdf14, Pdf15, Pdf16, or Pdf17.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void OutputIntentPolicy(PdfOutputIntentPolicy value, string paramName) {
        if (value != PdfOutputIntentPolicy.Unspecified &&
            value != PdfOutputIntentPolicy.SrgbIec6196621) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF output-intent policy must be Unspecified or SrgbIec6196621.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void TaggedStructureMode(PdfTaggedStructureMode value, string paramName) {
        if (value != PdfTaggedStructureMode.None &&
            value != PdfTaggedStructureMode.CatalogMarkers) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF tagged-structure mode must be None or CatalogMarkers.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void AssociatedFileRelationship(PdfAssociatedFileRelationship value, string paramName) {
        if (value != PdfAssociatedFileRelationship.Unspecified &&
            value != PdfAssociatedFileRelationship.Source &&
            value != PdfAssociatedFileRelationship.Data &&
            value != PdfAssociatedFileRelationship.Alternative &&
            value != PdfAssociatedFileRelationship.Supplement) {
            throw new System.ArgumentOutOfRangeException(paramName, "PDF associated-file relationship must be Unspecified, Source, Data, Alternative, or Supplement.");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void TextBaseline(PdfTextBaseline value, string paramName) {
        if (value != PdfTextBaseline.Normal &&
            value != PdfTextBaseline.Superscript &&
            value != PdfTextBaseline.Subscript) {
            throw new System.ArgumentException("PDF text baseline must be Normal, Superscript, or Subscript.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void TabLeaderStyle(PdfTabLeaderStyle value, string paramName) {
        if (value != PdfTabLeaderStyle.None &&
            value != PdfTabLeaderStyle.Dots &&
            value != PdfTabLeaderStyle.Hyphens &&
            value != PdfTabLeaderStyle.Underscores) {
            throw new System.ArgumentException("PDF tab leader style must be None, Dots, Hyphens, or Underscores.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void TabAlignment(PdfTabAlignment value, string paramName) {
        if (value != PdfTabAlignment.Left &&
            value != PdfTabAlignment.Center &&
            value != PdfTabAlignment.Right &&
            value != PdfTabAlignment.DecimalSeparator) {
            throw new System.ArgumentException("PDF tab alignment must be Left, Center, Right, or DecimalSeparator.", paramName);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void StandardFont(PdfStandardFont value, string paramName, string message) {
        if (value != PdfStandardFont.Helvetica &&
            value != PdfStandardFont.HelveticaOblique &&
            value != PdfStandardFont.HelveticaBold &&
            value != PdfStandardFont.HelveticaBoldOblique &&
            value != PdfStandardFont.TimesRoman &&
            value != PdfStandardFont.TimesItalic &&
            value != PdfStandardFont.TimesBold &&
            value != PdfStandardFont.TimesBoldItalic &&
            value != PdfStandardFont.Courier &&
            value != PdfStandardFont.CourierOblique &&
            value != PdfStandardFont.CourierBold &&
            value != PdfStandardFont.CourierBoldOblique) {
            throw new System.ArgumentOutOfRangeException(paramName, message);
        }
    }
}
