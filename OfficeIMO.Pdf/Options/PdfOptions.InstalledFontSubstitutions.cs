namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    // Called only after a system-loaded family was successfully registered. Explicit caller
    // registrations and rendering profiles do not pass through this automatic policy.
    private void ReportInstalledOfficeFontSubstitution(string requestedFamilies, PdfEmbeddedFontFamily family, bool namedFamily) {
        string[] candidates = EnumerateOfficeFontFamilyCandidates(requestedFamilies).ToArray();
        if (candidates.Any(candidate => string.Equals(candidate, family.FamilyName, StringComparison.OrdinalIgnoreCase))) return;
        string? source = candidates.FirstOrDefault(candidate =>
            OfficeIMO.Drawing.OfficeDocumentFontSubstitutions.Enumerate(candidate)
                .Contains(family.FamilyName, StringComparer.OrdinalIgnoreCase));
        if (source == null) return;

        PdfFontFamilySubstitutionImpact impact = OfficeIMO.Drawing.OfficeDocumentFontSubstitutions.IsIntendedCompatibleSubstitute(source, family.FamilyName)
            ? PdfFontFamilySubstitutionImpact.Compatible : PdfFontFamilySubstitutionImpact.LayoutSensitive;
        if (namedFamily && !TryGetFontFamilySubstitution(source, out _)) {
            RegisterFontFamilySubstitution(source, family.FamilyName, impact);
        }

        if (DiagnosticsReport == null ||
            !PdfOutputDiagnosticScope.Admit("installed-font-substitution|" + source + "|" + family.FamilyName, ref _reportedLayoutDiagnostics)) return;
        DiagnosticsReport.Add(new PdfConversionWarning(
            _diagnosticsConverter, "font-family-substitution", "font:" + source,
            "The source font family '" + source + "' is unavailable for embedding; generated text uses the installed substitute '" + family.FamilyName + "'. " +
            (impact == PdfFontFamilySubstitutionImpact.Compatible
                ? "This is an intended compatible substitute. " : string.Empty) +
            "Exact font-version metrics, including vertical metrics, wrapping, and pagination, have not been verified.",
            impact == PdfFontFamilySubstitutionImpact.Compatible ? PdfConversionWarningSeverity.Information : PdfConversionWarningSeverity.Warning,
            OfficeConversionLossKind.Approximation,
            details: new Dictionary<string, string> {
                ["fontFamily"] = source,
                ["resolvedFontFamily"] = family.FamilyName,
                ["substitutionReason"] = "FontUnavailable",
                ["substitutionImpact"] = impact.ToString(),
                ["substitutionPolicy"] = "OfficeDocumentFonts",
                ["compatibilityBasis"] = impact == PdfFontFamilySubstitutionImpact.Compatible ? "IntendedCompatibility" : "VisualSimilarity",
                ["fontMetricsVerification"] = "NotPerformed"
            }));
    }
}
