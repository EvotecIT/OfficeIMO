using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private Dictionary<string, PdfFontFamilySubstitution>? _fontFamilySubstitutions;

    /// <summary>
    /// Configured source-to-embedded font substitutions ordered by source family.
    /// </summary>
    public IReadOnlyList<PdfFontFamilySubstitution> FontFamilySubstitutions {
        get {
            if (_fontFamilySubstitutions == null || _fontFamilySubstitutions.Count == 0) {
                return Array.Empty<PdfFontFamilySubstitution>();
            }

            return new ReadOnlyCollection<PdfFontFamilySubstitution>(
                _fontFamilySubstitutions.Values
                    .OrderBy(static substitution => substitution.SourceFontFamily, StringComparer.OrdinalIgnoreCase)
                    .Select(static substitution => substitution.Clone())
                    .ToList());
        }
    }

    /// <summary>
    /// Registers a source document family as a planned substitute for an embedded named PDF family.
    /// </summary>
    /// <remarks>
    /// The target family must be registered through <see cref="RegisterNamedFontFamily"/> before conversion.
    /// Compatible substitutions are informational diagnostics; layout-sensitive substitutions remain warnings.
    /// </remarks>
    /// <param name="sourceFontFamily">Source family declared by the input document.</param>
    /// <param name="targetFontFamily">Registered embedded family used in generated PDF text.</param>
    /// <param name="impact">Expected layout impact of the substitution.</param>
    /// <returns>The current options instance.</returns>
    public PdfOptions RegisterFontFamilySubstitution(
        string sourceFontFamily,
        string targetFontFamily,
        PdfFontFamilySubstitutionImpact impact = PdfFontFamilySubstitutionImpact.LayoutSensitive) {
        var substitution = new PdfFontFamilySubstitution(sourceFontFamily, targetFontFamily, impact);
        string key = NormalizeOfficeFontFamilyKey(substitution.SourceFontFamily);
        (_fontFamilySubstitutions ??=
            new Dictionary<string, PdfFontFamilySubstitution>(StringComparer.OrdinalIgnoreCase))[key] = substitution;
        return this;
    }

    internal bool TryGetFontFamilySubstitution(
        string? sourceFontFamily,
        out PdfFontFamilySubstitution? substitution) {
        substitution = null;
        if (string.IsNullOrWhiteSpace(sourceFontFamily) || _fontFamilySubstitutions == null) {
            return false;
        }

        foreach (string candidate in EnumerateOfficeFontFamilyCandidates(sourceFontFamily!)) {
            if (_fontFamilySubstitutions.TryGetValue(
                    NormalizeOfficeFontFamilyKey(candidate),
                    out PdfFontFamilySubstitution? configured)) {
                substitution = configured;
                return true;
            }
        }

        return false;
    }

    internal bool TryResolveFontFamilySubstitution(
        string? sourceFontFamily,
        out PdfFontFamilySubstitution? substitution) {
        substitution = null;
        if (string.IsNullOrWhiteSpace(sourceFontFamily) || _fontFamilySubstitutions == null) {
            return false;
        }

        foreach (string candidate in EnumerateOfficeFontFamilyCandidates(sourceFontFamily!)) {
            if (_fontFamilySubstitutions.TryGetValue(
                    NormalizeOfficeFontFamilyKey(candidate),
                    out PdfFontFamilySubstitution? configured) &&
                TryGetNamedFontFamilyDirect(configured.TargetFontFamily, out _)) {
                substitution = configured;
                return true;
            }

            // A directly registered family earlier in an Office/CSS fallback list
            // takes precedence over substitutions configured for later candidates.
            if (TryGetNamedFontFamilyDirect(candidate, out _)) {
                return false;
            }
        }

        return false;
    }

    internal PdfConversionWarning CreateFontFamilySubstitutionWarning(
        string converter,
        string code,
        string source,
        string sourceFontFamily,
        PdfStandardFont? fallbackSlot,
        string? resolvedFontFamily,
        IReadOnlyDictionary<string, string>? additionalDetails = null) {
        var details = new Dictionary<string, string>(StringComparer.Ordinal);
        if (additionalDetails != null) {
            foreach (KeyValuePair<string, string> detail in additionalDetails) {
                details[detail.Key] = detail.Value;
            }
        }
        details["fontFamily"] = sourceFontFamily;

        PdfStandardFont? normalizedSlot = fallbackSlot.HasValue
            ? PdfStandardFontMapper.GetFontFamily(fallbackSlot.Value)
            : null;
        if (normalizedSlot.HasValue) {
            details["fallbackSlot"] = normalizedSlot.Value.ToString();
        }
        if (!string.IsNullOrWhiteSpace(resolvedFontFamily)) {
            details["resolvedFontFamily"] = resolvedFontFamily!;
        }

        if (TryResolveFontFamilySubstitution(sourceFontFamily, out PdfFontFamilySubstitution? substitution) &&
            substitution != null &&
            !string.IsNullOrWhiteSpace(resolvedFontFamily) &&
            string.Equals(
                 NormalizeOfficeFontFamilyKey(substitution.TargetFontFamily),
                 NormalizeOfficeFontFamilyKey(resolvedFontFamily!),
                 StringComparison.OrdinalIgnoreCase)) {
            details["substitutionImpact"] = substitution.Impact.ToString();
            details["plannedSubstitution"] = bool.TrueString;
            PdfConversionWarningSeverity severity =
                substitution.Impact == PdfFontFamilySubstitutionImpact.Compatible
                    ? PdfConversionWarningSeverity.Information
                    : PdfConversionWarningSeverity.Warning;
            string message = substitution.Impact == PdfFontFamilySubstitutionImpact.Compatible
                ? "The source font family '" + sourceFontFamily +
                  "' uses the configured compatible embedded substitute '" +
                  substitution.TargetFontFamily + "'."
                : "The source font family '" + sourceFontFamily +
                  "' uses the configured embedded substitute '" +
                  substitution.TargetFontFamily + "'.";
            return new PdfConversionWarning(
                converter,
                code,
                source,
                message,
                severity,
                details: details);
        }

        string fallbackDescription;
        if (!string.IsNullOrWhiteSpace(resolvedFontFamily)) {
            fallbackDescription = "the embedded family '" + resolvedFontFamily + "'" +
                (normalizedSlot.HasValue ? " through the logical " + normalizedSlot.Value + " PDF slot" : string.Empty);
        } else {
            fallbackDescription = normalizedSlot.HasValue
                ? "the mapped PDF family " + normalizedSlot.Value
                : "the configured default PDF family";
        }

        string fallbackCandidates = FormatEmbeddedFallbackFamilyNames();
        if (!string.IsNullOrWhiteSpace(fallbackCandidates)) {
            details["embeddedFallbackFamilies"] = fallbackCandidates;
        }
        string fallbackNote = string.IsNullOrWhiteSpace(fallbackCandidates)
            ? string.Empty
            : " Glyphs outside that family may use the embedded fallback families " + fallbackCandidates + ".";
        return new PdfConversionWarning(
            converter,
            code,
            source,
            "The source font family '" + sourceFontFamily +
            "' was unavailable or could not be embedded; generated text uses " +
            fallbackDescription + "." + fallbackNote,
            details: details);
    }

    private string FormatEmbeddedFallbackFamilyNames() {
        PdfEmbeddedFontFallbackSet? fallbackSet = EmbeddedFontFallbacksSnapshot;
        if (fallbackSet?.UsesNamedFontFamilies != true || fallbackSet.FontFamilyNames.Count == 0) {
            return string.Empty;
        }

        return string.Join(", ", fallbackSet.FontFamilyNames.Select(static family => "'" + family + "'"));
    }

    private static Dictionary<string, PdfFontFamilySubstitution>? CloneFontFamilySubstitutions(
        Dictionary<string, PdfFontFamilySubstitution>? substitutions) {
        if (substitutions == null) {
            return null;
        }

        var clone = new Dictionary<string, PdfFontFamilySubstitution>(StringComparer.OrdinalIgnoreCase);
        foreach (KeyValuePair<string, PdfFontFamilySubstitution> substitution in substitutions) {
            clone[substitution.Key] = substitution.Value.Clone();
        }

        return clone;
    }
}
