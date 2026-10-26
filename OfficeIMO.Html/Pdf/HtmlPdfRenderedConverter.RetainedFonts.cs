using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static readonly OfficeFontStyle[] RetainedDefaultStyles = {
        OfficeFontStyle.Regular, OfficeFontStyle.Bold,
        OfficeFontStyle.Italic, OfficeFontStyle.Bold | OfficeFontStyle.Italic
    };

    /// <summary>
    /// Makes an explicitly selected target default available to shared layout and paint.
    /// This imports existing snapshots only; the originating request still owns host access.
    /// A supplied CSS face keeps ownership of its family and style fallback.
    /// </summary>
    internal static IReadOnlyList<string> RetainConfiguredDefaultFonts(OfficeFontFaceCollection fonts,
        HtmlRenderOptions options, HtmlResourceSession resources, HtmlDiagnosticReport diagnostics,
        CancellationToken cancellationToken) {
        if (options is not HtmlToPdfOptions target) return Array.Empty<string>();
        var configuredFamilies = new List<string>();
        IReadOnlyDictionary<string, PdfCore.PdfEmbeddedFontFamily> named = target.PdfOptions.NamedFontFamilies;
        foreach (string candidate in EnumerateFamilies(options.DefaultFontFamily)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (HasRetainedFamily(fonts, candidate)) continue;
            if (target.PdfOptions.TryResolveNamedOfficeFontFamily(candidate, out string? registered)
                && registered != null
                && named.TryGetValue(registered, out PdfCore.PdfEmbeddedFontFamily? selected)) {
                configuredFamilies.Add(candidate);
                TryRetainDefaultFaces(candidate, FamilyFaces(selected), fonts, resources,
                    diagnostics, cancellationToken);
                continue;
            }

            PdfCore.PdfStandardFont slot = MapStandardFont(candidate);
            PdfCore.PdfEmbeddedFontFamily? configured = slot == PdfCore.PdfStandardFont.Helvetica
                ? target.FontFamily ?? target.LibrarySelectedDefaultFontFamily : null;
            if (configured != null) {
                configuredFamilies.Add(candidate);
                TryRetainDefaultFaces(candidate, FamilyFaces(configured), fonts, resources,
                    diagnostics, cancellationToken);
                continue;
            }

            // Standard registrations can be partial: retain only the supplied physical
            // faces, leaving the collection's existing style-fallback contract in charge.
            var faces = new List<(OfficeFontStyle Style, byte[] Data)>();
            foreach (OfficeFontStyle style in RetainedDefaultStyles) {
                PdfCore.PdfStandardFont variant = PdfCore.PdfStandardFontMapper.GetStyledFont(slot,
                    (style & OfficeFontStyle.Bold) != 0, (style & OfficeFontStyle.Italic) != 0);
                if (target.PdfOptions.TryGetEmbeddedStandardFont(variant, out PdfCore.PdfEmbeddedFont? embedded)
                    && embedded != null && !embedded.SyntheticOblique) faces.Add((style, embedded.DataSnapshot));
            }
            if (faces.Count > 0) {
                configuredFamilies.Add(candidate);
                TryRetainDefaultFaces(candidate, faces, fonts, resources, diagnostics, cancellationToken);
            }
        }
        // Failure does not authorize an unrelated installed replacement. The loader
        // preserves these owned names even when their bounded admission was rejected.
        return configuredFamilies;
    }

    /// <summary>
    /// Retains an operation-selected library default before its authoritative layout pass.
    /// Only configured default aliases are added; unrelated CSS families keep their own
    /// standard-slot and scoped-face behavior.
    /// </summary>
    internal static bool TryRetainSelectedDefaultFontFamily(PdfCore.PdfEmbeddedFontFamily family,
        OfficeFontFaceCollection fonts, HtmlToPdfOptions options, HtmlResourceSession resources,
        HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken) {
        bool retained = false;
        foreach (string candidate in EnumerateFamilies(options.DefaultFontFamily)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (HasRetainedFamily(fonts, candidate)
                || MapStandardFont(candidate) != PdfCore.PdfStandardFont.Helvetica) continue;
            retained |= TryRetainDefaultFaces(candidate, FamilyFaces(family), fonts,
                resources, diagnostics, cancellationToken);
        }
        return retained;
    }

    private static IEnumerable<(OfficeFontStyle Style, byte[] Data)> FamilyFaces(
        PdfCore.PdfEmbeddedFontFamily family) {
        yield return (OfficeFontStyle.Regular, family.RegularSnapshot);
        if (family.BoldSnapshot != null) yield return (OfficeFontStyle.Bold, family.BoldSnapshot);
        if (family.ItalicSnapshot != null) yield return (OfficeFontStyle.Italic, family.ItalicSnapshot);
        if (family.BoldItalicSnapshot != null)
            yield return (OfficeFontStyle.Bold | OfficeFontStyle.Italic, family.BoldItalicSnapshot);
    }

    private static bool HasRetainedFamily(OfficeFontFaceCollection fonts, string family) =>
        fonts.Faces.Any(face => string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase));

    private static bool TryRetainDefaultFaces(string family,
        IEnumerable<(OfficeFontStyle Style, byte[] Data)> faces, OfficeFontFaceCollection fonts,
        HtmlResourceSession resources, HtmlDiagnosticReport diagnostics,
        CancellationToken cancellationToken) {
        var staged = new OfficeFontFaceCollection {
            FontProgramProvider = fonts.FontProgramProvider,
            FontVariationResolver = fonts.FontVariationResolver
        };
        int acceptedBytes = 0;
        foreach ((OfficeFontStyle style, byte[] data) in faces) {
            cancellationToken.ThrowIfCancellationRequested();
            long remaining = Math.Min(resources.RemainingFontBytes - acceptedBytes,
                (long)int.MaxValue - acceptedBytes);
            if (data.LongLength > resources.MaxResourceBytes || remaining <= 0) {
                Report(data.LongLength > resources.MaxResourceBytes
                    ? HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded
                    : HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded,
                    "Selected font retention exceeds the operation's font-resource budget.");
                return false;
            }
            if (!staged.TryAddBounded(family, data, OfficeFontFaceDescriptor.FromStyle(style),
                    OfficeFontUnicodeRangeSet.All, (int)Math.Min(remaining, int.MaxValue),
                    out int decodedBytes, out string? error)) {
                Report(error?.IndexOf("limit", StringComparison.OrdinalIgnoreCase) >= 0
                    ? HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded
                    : HtmlRenderDiagnosticCodes.FontFormatUnsupported,
                    error ?? "The selected font could not be retained by the bounded font engine.");
                return false;
            }
            acceptedBytes = checked(acceptedBytes + decodedBytes);
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (staged.Faces.Count == 0) return false;
        // Admission is all-or-nothing. Failed staging cannot change the selected face
        // or leave measurement authoritative for a font that paint does not retain.
        resources.AcceptDecodedFontBytes(acceptedBytes);
        fonts.AddRangePreservingExisting(staged);
        return true;

        void Report(string code, string detail) => diagnostics.Add("OfficeIMO.Html.Renderer", code,
            "A selected default font could not be retained for shared layout and painting.",
            HtmlDiagnosticSeverity.Warning, family, detail);
    }
}
