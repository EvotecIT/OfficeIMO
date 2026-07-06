namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>
    /// Default installed symbol and emoji family candidates used by document converters for generated PDF text fallback runs.
    /// </summary>
    public const string DefaultDocumentSymbolAndEmojiFontFamilyFallback = "Segoe UI Symbol, Segoe UI Emoji, Noto Emoji, Noto Color Emoji, Noto Sans Symbols, Noto Sans Symbols 2, Symbola, DejaVu Sans, Arial Unicode MS, Arial";

    /// <summary>
    /// Applies OfficeIMO's built-in generated-text fallback groups without requiring callers to manually assign fallback font slots.
    /// </summary>
    /// <param name="features">Fallback groups to enable. The default enables document, monospace, symbol, and emoji fallbacks.</param>
    /// <returns>The current options for fluent chaining.</returns>
    public PdfOptions UseTextFallbacks(PdfTextFallbackFeatures features = PdfTextFallbackFeatures.Default) {
        if ((features & PdfTextFallbackFeatures.DocumentFont) != 0) {
            TryUseDefaultDocumentFontFallback(requireEmbeddedFont: false);
        }

        if ((features & PdfTextFallbackFeatures.MonospaceFont) != 0) {
            TryRegisterDefaultDocumentMonospaceFontFallback(requireEmbeddedFont: false);
        }

        if ((features & PdfTextFallbackFeatures.SymbolAndEmojiFonts) != 0) {
            TryRegisterEmbeddedFontFallbacksFromSystem(DefaultDocumentSymbolAndEmojiFontFamilyFallback);
        }

        return this;
    }

    /// <summary>
    /// Registers generated-text fallback fonts by resolving an Office-style comma/semicolon-separated system font family list.
    /// Callers do not need to choose PDF standard-font slots; OfficeIMO selects available slots and preserves explicit fallback sets.
    /// </summary>
    /// <param name="familyNames">System font family candidates, for example <c>Segoe UI Emoji, Noto Emoji, DejaVu Sans</c>.</param>
    /// <param name="maxFallbackFonts">Maximum number of installed fallback font families to register.</param>
    /// <returns>The current options for fluent chaining.</returns>
    public PdfOptions UseEmbeddedFontFallbacksFromSystem(string? familyNames, int maxFallbackFonts = 2) {
        TryRegisterEmbeddedFontFallbacksFromSystem(familyNames, maxFallbackFonts);
        return this;
    }

    /// <summary>
    /// Tries to register generated-text fallback fonts from installed system font families without requiring callers to choose PDF font slots.
    /// Existing explicit <see cref="EmbeddedFontFallbacks"/> are preserved.
    /// </summary>
    /// <param name="familyNames">System font family candidates, for example <c>Segoe UI Emoji, Noto Emoji, DejaVu Sans</c>.</param>
    /// <param name="maxFallbackFonts">Maximum number of installed fallback font families to register.</param>
    /// <returns>True when fallback fonts are already configured or at least one installed fallback was registered.</returns>
    public bool TryRegisterEmbeddedFontFallbacksFromSystem(string? familyNames, int maxFallbackFonts = 2) {
        if (maxFallbackFonts <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxFallbackFonts), "Maximum fallback font count must be positive.");
        }

        if (_embeddedFontFallbacks != null) {
            return true;
        }

        if (string.IsNullOrWhiteSpace(familyNames)) {
            return false;
        }

        var candidates = new List<PdfEmbeddedFontFallbackCandidate>();
        var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string familyName in EnumerateOfficeFontFamilyCandidates(familyNames!)) {
            if (candidates.Count == maxFallbackFonts) {
                break;
            }

            if (!registeredFamilies.Add(familyName)) {
                continue;
            }

            if (PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfEmbeddedFontFamily? family) &&
                family != null) {
                candidates.Add(new PdfEmbeddedFontFallbackCandidate(family.FamilyName, family.Regular));
            }
        }

        if (candidates.Count == 0) {
            return false;
        }

        PdfStandardFont[] slots = GetAvailableEmbeddedFallbackFontSlots(candidates.Count, Array.Empty<PdfStandardFont>()).ToArray();
        if (slots.Length == 0) {
            return false;
        }

        if (slots.Length < candidates.Count) {
            candidates = SelectPreferredEmbeddedFallbackCandidates(candidates, slots.Length);
        }

        RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(candidates, slots));
        return true;
    }

    private static List<PdfEmbeddedFontFallbackCandidate> SelectPreferredEmbeddedFallbackCandidates(
        IReadOnlyList<PdfEmbeddedFontFallbackCandidate> candidates,
        int slotCount) {
        var selected = new List<PdfEmbeddedFontFallbackCandidate>();
        if (slotCount <= 0) {
            return selected;
        }

        foreach (PdfEmbeddedFontFallbackCandidate candidate in candidates) {
            if (selected.Count == slotCount) {
                return selected;
            }

            if (IsEmojiFallbackCandidate(candidate)) {
                selected.Add(candidate);
            }
        }

        foreach (PdfEmbeddedFontFallbackCandidate candidate in candidates) {
            if (selected.Count == slotCount) {
                return selected;
            }

            if (!selected.Contains(candidate)) {
                selected.Add(candidate);
            }
        }

        return selected;
    }

    private static bool IsEmojiFallbackCandidate(PdfEmbeddedFontFallbackCandidate candidate) =>
        System.Globalization.CultureInfo.InvariantCulture.CompareInfo.IndexOf(
            candidate.FontName,
            "Emoji",
            System.Globalization.CompareOptions.IgnoreCase) >= 0;
}
