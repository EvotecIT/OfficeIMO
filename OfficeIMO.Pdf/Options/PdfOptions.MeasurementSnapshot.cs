namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    /// <summary>Creates an internal text-only measurement view: assets are borrowed, while font usage and diagnostic caches are independent.</summary>
    /// <remarks>Never expose this view for general editing or serialization. It must not mutate borrowed attachments, styles, or font definitions.</remarks>
    internal PdfCreateOptions CreateTextMeasurementSnapshot() {
        var view = (PdfCreateOptions)MemberwiseClone();
        view._themeDefaults = null;
        view._themeDeferredTableStyle = null;
        // Font definitions and their immutable owned bytes can be borrowed. Dictionary
        // membership must be private because automatic fallback may register a family.
        view._embeddedFonts = _embeddedFonts?.CopyWithSameEntries();
        // A deferred default family resolves into whichever options ask first, so the view keeps its own record.
        view._pendingDefaultFontFamilies = ClonePendingDefaultFontFamilies();
        view._defaultEmbeddedFontFamilies = _defaultEmbeddedFontFamilies == null ? null : new HashSet<PdfStandardFont>(_defaultEmbeddedFontFamilies);
        view._namedFontFamilies = _namedFontFamilies?.CopyWithSameEntries();
        view._namedFallbackPlanBindings = _namedFallbackPlanBindings == null ? null : new Dictionary<object, PdfEmbeddedFontFallbackCandidate>(_namedFallbackPlanBindings);
        view._embeddedFontPrograms = null;
        view._embeddedOpenTypeCffFontPrograms = null;
        view._namedFontPrograms = null;
        view._namedOpenTypeCffFontPrograms = null;
        view._namedFontProgramFailures = null;
        view._embeddedFontProgramFailures = null;
        view._reportedEmbeddedFontProgramFailures = null;
        view._reportedTextShapingDiagnostics = null;
        view._reportedLayoutDiagnostics = null;
        view._providerShapedTextRuns = null;
        view._automaticLatinShapedTextRun = null;
        view._recordProviderShapedTextRunDelegate = null;
        view._diagnosticsReport = null;
        view._suppressDiagnostics = true;
        view._fallbackOwnedFontFamilies = _fallbackOwnedFontFamilies == null ? null : new HashSet<PdfStandardFont>(_fallbackOwnedFontFamilies);
        view._usedEmbeddedFallbackFontSlots = CloneUsedEmbeddedFallbackFontSlots();
        view._automaticFallbackOwnedNamedFamilyKeys = _automaticFallbackOwnedNamedFamilyKeys == null ? null : new HashSet<string>(_automaticFallbackOwnedNamedFamilyKeys, StringComparer.Ordinal);
        view._renderingProfileOwnedNamedFamilyNames = _renderingProfileOwnedNamedFamilyNames == null ? null : new HashSet<string>(_renderingProfileOwnedNamedFamilyNames, StringComparer.OrdinalIgnoreCase);
        return view;
    }

    /// <summary>Checks selected immutable font-definition identities without hashing or copying their payloads.</summary>
    internal bool HasSameMeasurementFonts(PdfCreateOptions source) {
        if (_fontConfigurationAssignmentVersion != source._fontConfigurationAssignmentVersion ||
            _defaultFontSizeAssignmentVersion != source._defaultFontSizeAssignmentVersion ||
            (_embeddedFonts?.Count ?? 0) != (source._embeddedFonts?.Count ?? 0) ||
            (_namedFontFamilies?.Count ?? 0) != (source._namedFontFamilies?.Count ?? 0)) return false;
        bool sameEmbedded = (_embeddedFonts?.Count ?? 0) == 0
            || PdfFontMembershipMap<PdfStandardFont, PdfEmbeddedFont>.HaveSameEntries(_embeddedFonts, source._embeddedFonts);
        bool sameNamed = (_namedFontFamilies?.Count ?? 0) == 0
            || PdfFontMembershipMap<string, PdfEmbeddedFontFamily>.HaveSameEntries(_namedFontFamilies, source._namedFontFamilies);
        if (sameEmbedded && sameNamed) return true;
        if (source._embeddedFonts != null) foreach (var entry in source._embeddedFonts)
            if (_embeddedFonts == null || !_embeddedFonts.TryGetValue(entry.Key, out var definition) || !ReferenceEquals(definition, entry.Value)) return false;
        if (source._namedFontFamilies != null) foreach (var entry in source._namedFontFamilies)
            if (_namedFontFamilies == null || !_namedFontFamilies.TryGetValue(entry.Key, out var family) || !ReferenceEquals(family, entry.Value)) return false;
        return true;
    }
}
