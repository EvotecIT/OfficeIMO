namespace OfficeIMO.Pdf;

/// <summary>Preview of text, image placements, and annotations that intersect requested redaction rectangles.</summary>
public sealed class PdfRedactionPlan {
    internal PdfRedactionPlan(
        PdfDocumentPreflight preflight,
        IReadOnlyList<PdfRedactionArea> areas,
        IReadOnlyList<PdfRedactionMatch> matches,
        IReadOnlyList<PdfDiagnosticFinding> findings,
        IReadOnlyList<string>? searchCriteria,
        string sourceSha256,
        IReadOnlyList<string>? pageIdentities = null) {
        Preflight = preflight;
        Areas = areas;
        Matches = matches;
        Findings = findings;
        SearchCriteria = searchCriteria ?? Array.Empty<string>();
        SourceSha256 = sourceSha256;
        PageIdentities = pageIdentities ?? Array.Empty<string>();
    }

    /// <summary>Preflight result used while creating the plan.</summary>
    public PdfDocumentPreflight Preflight { get; }

    /// <summary>Requested redaction areas.</summary>
    public IReadOnlyList<PdfRedactionArea> Areas { get; }

    /// <summary>Text blocks, image placements, and annotations intersecting the requested areas.</summary>
    public IReadOnlyList<PdfRedactionMatch> Matches { get; }

    /// <summary>Diagnostics and warnings for the plan.</summary>
    public IReadOnlyList<PdfDiagnosticFinding> Findings { get; }

    /// <summary>Stable descriptions of literal, regex, logical-kind, or form-field criteria used to derive the areas.</summary>
    public IReadOnlyList<string> SearchCriteria { get; }

    /// <summary>SHA-256 fingerprint of the exact PDF bytes inspected while creating this plan.</summary>
    public string SourceSha256 { get; }

    internal IReadOnlyList<string> PageIdentities { get; }

    /// <summary>True when the source was inspectable and the plan contains no blocking findings.</summary>
    public bool IsReviewable =>
        Preflight.CanReadLogicalObjects &&
        Findings.All(static finding => finding.Severity != PdfDiagnosticSeverity.Error);

    /// <summary>True when the plan areas were derived from explicit search criteria.</summary>
    public bool IsSearchDriven => SearchCriteria.Count > 0;

    /// <summary>True when at least one match was found.</summary>
    public bool HasMatches => Matches.Count > 0;

    internal bool MatchesSource(byte[] pdf) =>
        string.Equals(SourceSha256, ComputeSourceSha256(pdf), StringComparison.Ordinal);

    internal static string ComputeSourceSha256(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
#if NET6_0_OR_GREATER
        return Convert.ToBase64String(System.Security.Cryptography.SHA256.HashData(pdf));
#else
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(pdf));
#endif
    }

    internal static IReadOnlyList<string> CapturePageIdentities(
        PdfReadDocument document,
        IReadOnlyList<PdfRedactionArea> reviewedAreas) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(reviewedAreas, nameof(reviewedAreas));
        var identities = new string[document.Pages.Count];
        for (int i = 0; i < document.Pages.Count; i++) {
            PdfReadPage page = document.Pages[i];
            PdfPageGeometry geometry = page.GetGeometry();
            int pageNumber = i + 1;
            PdfRedactionArea[] pageAreas = reviewedAreas
                .Where(area => area.PageNumber == pageNumber)
                .ToArray();
            var identity = new System.Text.StringBuilder();
            identity.Append(string.Join("|", new[] {
                page.GetRotationDegrees().ToString(System.Globalization.CultureInfo.InvariantCulture),
                FormatPageBoxIdentity(geometry.MediaBox),
                FormatPageBoxIdentity(geometry.CropBox),
                geometry.UserUnit?.ToString("R", System.Globalization.CultureInfo.InvariantCulture) ?? "null"
            }));
            AppendUnredactedTextIdentity(identity, page, pageAreas);
            AppendUnredactedImageIdentity(identity, document, page, pageNumber, pageAreas);
            AppendUnredactedAnnotationIdentity(identity, page, pageAreas);
            AppendUnredactedLinkIdentity(identity, page, pageAreas);
            identities[i] = ComputeIdentityHash(identity.ToString());
        }

        return identities;
    }

    private static void AppendUnredactedTextIdentity(
        System.Text.StringBuilder identity,
        PdfReadPage page,
        IReadOnlyList<PdfRedactionArea> pageAreas) {
        IReadOnlyList<PdfTextSpan> spans = page.GetTextSpans();
        for (int i = 0; i < spans.Count; i++) {
            PdfTextSpan span = spans[i];
            double x = Math.Min(span.X, span.X + span.Advance);
            double width = Math.Max(1D, Math.Abs(span.Advance));
            double height = Math.Max(1D, span.FontSize);
            double y = span.Y - height;
            if (IntersectsReviewedArea(pageAreas, x, y, width, height)) continue;
            identity.Append("|T:")
                .Append(span.Text.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(':').Append(span.Text)
                .Append(':').Append(FormatIdentityNumber(span.X))
                .Append(',').Append(FormatIdentityNumber(span.Y))
                .Append(',').Append(FormatIdentityNumber(span.Advance))
                .Append(',').Append(FormatIdentityNumber(span.FontSize))
                .Append(',').Append(FormatIdentityNumber(span.RotationDegrees))
                .Append(',').Append(span.IsVisible ? '1' : '0');
        }
    }

    private static void AppendUnredactedImageIdentity(
        System.Text.StringBuilder identity,
        PdfReadDocument document,
        PdfReadPage page,
        int pageNumber,
        IReadOnlyList<PdfRedactionArea> pageAreas) {
        IReadOnlyList<PdfImagePlacement> placements = page.GetImagePlacements();
        for (int i = 0; i < placements.Count; i++) {
            PdfImagePlacement placement = placements[i];
            if (IntersectsReviewedArea(pageAreas, placement.X, placement.Y, placement.Width, placement.Height)) continue;
            byte[]? imageBytes = null;
            if (placement.ObjectNumber > 0 &&
                document.Objects.TryGetValue(placement.ObjectNumber, out PdfIndirectObject? indirect) &&
                indirect.Value is PdfStream stream) {
                imageBytes = stream.Data;
            } else if (placement.InlineImageStream is PdfStream inlineStream) {
                imageBytes = inlineStream.Data;
            }

            identity.Append("|I:")
                .Append(pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(':').Append(FormatIdentityNumber(placement.A))
                .Append(',').Append(FormatIdentityNumber(placement.B))
                .Append(',').Append(FormatIdentityNumber(placement.C))
                .Append(',').Append(FormatIdentityNumber(placement.D))
                .Append(',').Append(FormatIdentityNumber(placement.E))
                .Append(',').Append(FormatIdentityNumber(placement.F))
                .Append(':').Append(imageBytes is null ? "none" : ComputeIdentityHash(imageBytes));
        }
    }

    private static void AppendUnredactedAnnotationIdentity(
        System.Text.StringBuilder identity,
        PdfReadPage page,
        IReadOnlyList<PdfRedactionArea> pageAreas) {
        IReadOnlyList<PdfAnnotation> annotations = page.GetAnnotations();
        for (int i = 0; i < annotations.Count; i++) {
            PdfAnnotation annotation = annotations[i];
            if (IntersectsReviewedArea(pageAreas, annotation.X1, annotation.Y1, annotation.Width, annotation.Height)) continue;

            identity.Append("|A:");
            AppendIdentityString(identity, annotation.Subtype);
            identity.Append(':').Append(FormatIdentityNumber(annotation.X1))
                .Append(',').Append(FormatIdentityNumber(annotation.Y1))
                .Append(',').Append(FormatIdentityNumber(annotation.X2))
                .Append(',').Append(FormatIdentityNumber(annotation.Y2));
            AppendIdentityString(identity, annotation.Contents);
            AppendIdentityString(identity, annotation.Name);
            AppendIdentityString(identity, annotation.Title);
            AppendIdentityString(identity, annotation.ActionType);
            identity.Append(':').Append(annotation.Flags?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "null")
                .Append(':').Append(annotation.HasNormalAppearance ? '1' : '0');
            AppendIdentityNumbers(identity, annotation.Color);
            AppendIdentityNumbers(identity, annotation.InteriorColor);
            AppendIdentityNumbers(identity, annotation.QuadPoints);
            AppendIdentityNumbers(identity, annotation.LineCoordinates);
            AppendIdentityNumbers(identity, annotation.Vertices);
            for (int pathIndex = 0; pathIndex < annotation.InkList.Count; pathIndex++) {
                AppendIdentityNumbers(identity, annotation.InkList[pathIndex]);
            }
            for (int actionIndex = 0; actionIndex < annotation.AdditionalActions.Count; actionIndex++) {
                PdfAnnotationAdditionalAction action = annotation.AdditionalActions[actionIndex];
                AppendIdentityString(identity, action.TriggerName);
                AppendIdentityString(identity, action.ActionType);
            }
            for (int actionIndex = 0; actionIndex < annotation.ChainedActions.Count; actionIndex++) {
                PdfAnnotationChainedAction action = annotation.ChainedActions[actionIndex];
                AppendIdentityString(identity, action.SourceName);
                AppendIdentityString(identity, action.ActionPath);
                AppendIdentityString(identity, action.ActionType);
            }
            if (annotation.Review != null) {
                AppendIdentityString(identity, annotation.Review.ReplyType);
                AppendIdentityString(identity, annotation.Review.State);
                AppendIdentityString(identity, annotation.Review.StateModel);
                AppendIdentityString(identity, annotation.Review.Subject);
                AppendIdentityString(identity, annotation.Review.Intent);
            }
        }
    }

    private static void AppendUnredactedLinkIdentity(
        System.Text.StringBuilder identity,
        PdfReadPage page,
        IReadOnlyList<PdfRedactionArea> pageAreas) {
        IReadOnlyList<PdfLinkAnnotation> links = page.GetLinkAnnotations();
        for (int i = 0; i < links.Count; i++) {
            PdfLinkAnnotation link = links[i];
            if (IntersectsReviewedArea(pageAreas, link.X1, link.Y1, link.Width, link.Height)) continue;

            identity.Append("|L:")
                .Append(FormatIdentityNumber(link.X1)).Append(',')
                .Append(FormatIdentityNumber(link.Y1)).Append(',')
                .Append(FormatIdentityNumber(link.X2)).Append(',')
                .Append(FormatIdentityNumber(link.Y2));
            AppendIdentityString(identity, link.Contents);
            AppendIdentityString(identity, link.Uri);
            AppendIdentityString(identity, link.DestinationName);
            AppendIdentityString(identity, link.NamedAction);
            AppendIdentityString(identity, link.RemoteFile);
            AppendIdentityString(identity, link.RemoteDestinationName);
            identity.Append(':').Append(link.DestinationPageNumber?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "null")
                .Append(':').Append(link.DestinationMode?.ToString() ?? "null")
                .Append(':').Append(link.RemoteDestinationPageNumber?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "null")
                .Append(':').Append(link.RemoteDestinationMode?.ToString() ?? "null");
            AppendIdentityNullableNumber(identity, link.DestinationLeft);
            AppendIdentityNullableNumber(identity, link.DestinationTop);
            AppendIdentityNullableNumber(identity, link.DestinationBottom);
            AppendIdentityNullableNumber(identity, link.DestinationRight);
            AppendIdentityNullableNumber(identity, link.RemoteDestinationLeft);
            AppendIdentityNullableNumber(identity, link.RemoteDestinationTop);
            AppendIdentityNullableNumber(identity, link.RemoteDestinationBottom);
            AppendIdentityNullableNumber(identity, link.RemoteDestinationRight);
        }
    }

    private static void AppendIdentityNullableNumber(System.Text.StringBuilder identity, double? value) =>
        identity.Append(':').Append(value.HasValue ? FormatIdentityNumber(value.Value) : "null");

    private static void AppendIdentityString(System.Text.StringBuilder identity, string? value) {
        if (value == null) {
            identity.Append(":null");
            return;
        }
        identity.Append(':')
            .Append(value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append(':').Append(value);
    }

    private static void AppendIdentityNumbers(
        System.Text.StringBuilder identity,
        IReadOnlyList<double> values) {
        identity.Append(':').Append(values.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        for (int i = 0; i < values.Count; i++) {
            identity.Append(',').Append(FormatIdentityNumber(values[i]));
        }
    }

    private static bool IntersectsReviewedArea(
        IReadOnlyList<PdfRedactionArea> areas,
        double x,
        double y,
        double width,
        double height) {
        for (int i = 0; i < areas.Count; i++) {
            PdfRedactionArea area = areas[i];
            if (x < area.X + area.Width && x + width > area.X &&
                y < area.Y + area.Height && y + height > area.Y) {
                return true;
            }
        }
        return false;
    }

    private static string FormatIdentityNumber(double value) =>
        value.ToString("R", System.Globalization.CultureInfo.InvariantCulture);

    private static string ComputeIdentityHash(string value) =>
        ComputeIdentityHash(System.Text.Encoding.UTF8.GetBytes(value));

    private static string ComputeIdentityHash(byte[] value) {
#if NET6_0_OR_GREATER
        return Convert.ToBase64String(System.Security.Cryptography.SHA256.HashData(value));
#else
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(value));
#endif
    }

    private static string FormatPageBoxIdentity(PdfPageBox? box) {
        if (box == null) {
            return "null";
        }

        return string.Join(",", new[] {
            box.Left.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            box.Bottom.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            box.Right.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            box.Top.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
        });
    }
}
