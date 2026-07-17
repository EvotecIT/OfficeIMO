namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>
    /// Default installed multilingual family candidates used by document converters for CJK,
    /// Arabic, and other non-Latin generated PDF text.
    /// </summary>
    public const string DefaultDocumentMultilingualFontFamilyFallback = "Arial Unicode MS, Noto Sans CJK JP, Yu Gothic, PingFang SC, Microsoft YaHei, Noto Sans CJK SC, Meiryo, Hiragino Sans GB, Microsoft JhengHei, Noto Sans CJK TC, Malgun Gothic, Apple SD Gothic Neo, Noto Sans CJK KR, MS Gothic, SimSun, Noto Sans JP, Noto Sans SC, Noto Sans TC, Noto Sans KR, Noto Sans Arabic, Noto Naskh Arabic, Arabic Typesetting, Traditional Arabic, Nirmala UI, Microsoft Uighur, DejaVu Sans";

    /// <summary>
    /// Default installed symbol and emoji family candidates used by document converters for generated PDF text fallback runs.
    /// </summary>
    public const string DefaultDocumentSymbolAndEmojiFontFamilyFallback = "Segoe UI Symbol, Apple Symbols, DejaVu Sans, Noto Sans Symbols 2, Noto Sans Symbols, Symbola, Arial Unicode MS, Segoe UI Emoji, Noto Emoji, Noto Color Emoji, Arial";

    /// <summary>
    /// Applies OfficeIMO's built-in generated-text fallback groups without requiring callers to manually assign fallback font slots.
    /// </summary>
    /// <param name="features">Fallback groups to enable. The default enables document, monospace, symbol, and emoji fallbacks.</param>
    /// <returns>The current options for fluent chaining.</returns>
    public PdfOptions UseTextFallbacks(PdfTextFallbackFeatures features = PdfTextFallbackFeatures.Default) {
        return UseTextFallbacks(features, Array.Empty<PdfStandardFont>(), allowSystemFontEmbedding: true);
    }

    internal PdfOptions UseTextFallbacks(
        PdfTextFallbackFeatures features,
        IEnumerable<PdfStandardFont> reservedFontSlots,
        bool allowSystemFontEmbedding) {
        Guard.NotNull(reservedFontSlots, nameof(reservedFontSlots));
        if (!allowSystemFontEmbedding || features == PdfTextFallbackFeatures.None) {
            return this;
        }

        var reservedSlots = new HashSet<PdfStandardFont>();
        foreach (PdfStandardFont slot in reservedFontSlots) {
            AddRegisteredFontFamilySlot(reservedSlots, slot);
        }

        if ((features & PdfTextFallbackFeatures.DocumentFont) != 0) {
            PdfStandardFont documentSlot = PdfStandardFontMapper.GetFontFamily(DefaultFont);
            if (!reservedSlots.Contains(documentSlot)) {
                TryUseDefaultDocumentFontFallback(requireEmbeddedFont: false);
            }
        }

        PdfTextFallbackFeatures runFallbacks = features &
            (PdfTextFallbackFeatures.MultilingualFonts | PdfTextFallbackFeatures.SymbolAndEmojiFonts);
        if (runFallbacks == PdfTextFallbackFeatures.SymbolAndEmojiFonts) {
            AddRegisteredFontFamilySlot(reservedSlots, PdfStandardFont.TimesRoman);
            TryRegisterEmbeddedFontFallbacksFromSystem(
                DefaultDocumentSymbolAndEmojiFontFamilyFallback,
                reservedFontSlots: reservedSlots);
        } else if (runFallbacks != PdfTextFallbackFeatures.None) {
            TryRegisterRunFallbacksFromSystem(runFallbacks, reservedSlots);
        }

        if ((features & PdfTextFallbackFeatures.MonospaceFont) != 0) {
            if (!reservedSlots.Contains(PdfStandardFont.Courier)) {
                TryRegisterDefaultDocumentMonospaceFontFallback(requireEmbeddedFont: false);
            }
        }

        return this;
    }

    private bool TryRegisterRunFallbacksFromSystem(
        PdfTextFallbackFeatures features,
        IEnumerable<PdfStandardFont> reservedFontSlots) {
        if (_embeddedFontFallbacks != null) return true;

        bool multilingual = (features & PdfTextFallbackFeatures.MultilingualFonts) != 0;
        bool symbols = (features & PdfTextFallbackFeatures.SymbolAndEmojiFonts) != 0;
        var reservedSlots = new HashSet<PdfStandardFont>();
        foreach (PdfStandardFont slot in reservedFontSlots) AddRegisteredFontFamilySlot(reservedSlots, slot);
        AddRegisteredFontFamilySlot(reservedSlots, PdfStandardFont.TimesRoman);
        var candidates = new List<PdfEmbeddedFontFallbackCandidate>();
        var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (multilingual) AddInstalledRunFallbackCandidates(
            DefaultDocumentMultilingualFontFamilyFallback,
            multilingual && symbols ? 1 : 2,
            candidates,
            registeredFamilies);
        if (symbols) AddInstalledRunFallbackCandidates(
            DefaultDocumentSymbolAndEmojiFontFamilyFallback,
            multilingual && symbols ? 1 : 2,
            candidates,
            registeredFamilies);
        if (candidates.Count < 2) {
            string remaining = multilingual
                ? DefaultDocumentMultilingualFontFamilyFallback
                : DefaultDocumentSymbolAndEmojiFontFamilyFallback;
            AddInstalledRunFallbackCandidates(remaining, 2 - candidates.Count, candidates, registeredFamilies);
        }

        PdfStandardFont[] slots = GetAvailableEmbeddedFallbackFontSlots(candidates.Count, reservedSlots).ToArray();
        if (slots.Length == 0 || candidates.Count == 0) return false;
        if (slots.Length < candidates.Count) candidates = SelectPreferredEmbeddedFallbackCandidates(candidates, slots.Length);
        RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(candidates, slots));
        return true;
    }

    private static void AddInstalledRunFallbackCandidates(
        string familyNames,
        int count,
        List<PdfEmbeddedFontFallbackCandidate> candidates,
        HashSet<string> registeredFamilies) {
        foreach (string familyName in EnumerateOfficeFontFamilyCandidates(familyNames)) {
            if (count == 0) return;
            if (!registeredFamilies.Add(familyName)) continue;
            if (!PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfEmbeddedFontFamily? family) || family == null) continue;
            candidates.Add(new PdfEmbeddedFontFallbackCandidate(family.FamilyName, family.Regular));
            count--;
        }
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
        return TryRegisterEmbeddedFontFallbacksFromSystem(familyNames, maxFallbackFonts, Array.Empty<PdfStandardFont>());
    }

    internal bool TryRegisterEmbeddedFontFallbacksFromSystem(
        string? familyNames,
        int maxFallbackFonts = 2,
        IEnumerable<PdfStandardFont>? reservedFontSlots = null) {
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

        PdfStandardFont[] slots = GetAvailableEmbeddedFallbackFontSlots(candidates.Count, reservedFontSlots ?? Array.Empty<PdfStandardFont>()).ToArray();
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

            if (!IsEmojiFallbackCandidate(candidate)) {
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
