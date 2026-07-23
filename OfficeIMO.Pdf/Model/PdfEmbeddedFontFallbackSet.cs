namespace OfficeIMO.Pdf;

/// <summary>
/// Pairs embedded fallback candidates with generated PDF font slots used for rendering planned fallback runs.
/// </summary>
public sealed class PdfEmbeddedFontFallbackSet {
    private readonly List<PdfEmbeddedFontFallbackCandidate> _candidates;
    private readonly List<PdfStandardFont> _fontSlots;
    private readonly List<string> _fontFamilyNames;

    /// <summary>
    /// Creates a reusable fallback set whose candidates are registered as named embedded families.
    /// Named fallback families do not consume the Helvetica, Times, or Courier compatibility slots,
    /// so they can coexist with document fonts that already use all three slots.
    /// </summary>
    /// <param name="candidates">Fallback font candidates in priority order. Candidate names must be unique.</param>
    public PdfEmbeddedFontFallbackSet(IEnumerable<PdfEmbeddedFontFallbackCandidate> candidates) {
        Guard.NotNull(candidates, nameof(candidates));

        _candidates = candidates.Select(CloneCandidate).ToList();
        _fontSlots = new List<PdfStandardFont>();
        _fontFamilyNames = _candidates.Select(candidate => candidate.FontName.Trim()).ToList();

        ValidateCandidates();
        if (_fontFamilyNames.Distinct(StringComparer.OrdinalIgnoreCase).Count() != _fontFamilyNames.Count) {
            throw new ArgumentException("Named embedded font fallback candidates must use distinct font names.", nameof(candidates));
        }
    }

    /// <summary>
    /// Creates a reusable fallback set from prioritized candidates and matching generated font slots.
    /// </summary>
    /// <param name="candidates">Fallback font candidates in priority order.</param>
    /// <param name="fontSlots">Generated standard-font family slots ordered the same way as <paramref name="candidates"/>.</param>
    public PdfEmbeddedFontFallbackSet(
        IEnumerable<PdfEmbeddedFontFallbackCandidate> candidates,
        IEnumerable<PdfStandardFont> fontSlots) {
        Guard.NotNull(candidates, nameof(candidates));
        Guard.NotNull(fontSlots, nameof(fontSlots));

        _candidates = candidates.Select(CloneCandidate).ToList();
        _fontSlots = fontSlots.Select(NormalizeFontSlot).ToList();
        _fontFamilyNames = new List<string>();

        ValidateCandidates();

        if (_candidates.Count != _fontSlots.Count) {
            throw new ArgumentException("Embedded font fallback candidates and font slots must have the same number of entries.", nameof(fontSlots));
        }

        var usedSlots = new HashSet<PdfStandardFont>();
        foreach (PdfStandardFont slot in _fontSlots) {
            if (!usedSlots.Add(slot)) {
                throw new ArgumentException("Embedded font fallback font slots must use distinct generated standard-font families.", nameof(fontSlots));
            }
        }
    }

    /// <summary>Fallback font candidates in priority order.</summary>
    public IReadOnlyList<PdfEmbeddedFontFallbackCandidate> Candidates => _candidates.AsReadOnly();

    /// <summary>Generated standard-font family slots ordered the same way as <see cref="Candidates"/>.</summary>
    public IReadOnlyList<PdfStandardFont> FontSlots => _fontSlots.AsReadOnly();

    /// <summary>
    /// Named embedded font families ordered the same way as <see cref="Candidates"/>.
    /// This is populated when the fallback set was created without compatibility slots.
    /// </summary>
    public IReadOnlyList<string> FontFamilyNames => _fontFamilyNames.AsReadOnly();

    /// <summary>True when fallback runs use named embedded font families instead of compatibility slots.</summary>
    public bool UsesNamedFontFamilies => _fontFamilyNames.Count != 0;

    internal PdfEmbeddedFontFallbackSet Clone() => UsesNamedFontFamilies
        ? new PdfEmbeddedFontFallbackSet(_candidates)
        : new PdfEmbeddedFontFallbackSet(_candidates, _fontSlots);

    /// <summary>
    /// Registers every fallback candidate into its generated font family slot, including bold and italic variants.
    /// </summary>
    /// <param name="options">PDF options to configure.</param>
    /// <returns>The supplied options for fluent chaining.</returns>
    public PdfOptions RegisterFonts(PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        if (UsesNamedFontFamilies) {
            foreach (PdfEmbeddedFontFallbackCandidate candidate in _candidates) {
                options.RegisterNamedFontFamily(
                    new PdfEmbeddedFontFamily(candidate.FontName, candidate.DataSnapshot));
            }

            return options;
        }

        RegisterFonts(options, _fontSlots);
        return options;
    }

    internal PdfOptions RegisterFonts(PdfOptions options, IReadOnlyList<PdfStandardFont> fontSlots) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(fontSlots, nameof(fontSlots));
        if (fontSlots.Count != _candidates.Count) {
            throw new ArgumentException("Embedded font fallback candidates and font slots must have the same number of entries.", nameof(fontSlots));
        }

        for (int index = 0; index < _candidates.Count; index++) {
            PdfEmbeddedFontFallbackCandidate candidate = _candidates[index];
            options.RegisterFontFamily(
                NormalizeFontSlot(fontSlots[index]),
                new PdfEmbeddedFontFamily(candidate.FontName, candidate.DataSnapshot));
        }

        return options;
    }

    internal PdfOptions RegisterFonts(PdfOptions options, IReadOnlyDictionary<int, PdfStandardFont> fontSlots) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(fontSlots, nameof(fontSlots));

        foreach (KeyValuePair<int, PdfStandardFont> entry in fontSlots) {
            int candidateIndex = entry.Key;
            if (candidateIndex < 0 || candidateIndex >= _candidates.Count) {
                throw new ArgumentException("Embedded font fallback font slots contain an unknown candidate index.", nameof(fontSlots));
            }

            PdfEmbeddedFontFallbackCandidate candidate = _candidates[candidateIndex];
            options.RegisterFontFamily(
                NormalizeFontSlot(entry.Value),
                new PdfEmbeddedFontFamily(candidate.FontName, candidate.DataSnapshot));
        }

        return options;
    }

    /// <summary>
    /// Plans fallback coverage for text using this set's prioritized candidates.
    /// </summary>
    /// <param name="text">Text to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="shapingMode">Text shaping mode to use when checking fallback font coverage.</param>
    /// <returns>A fallback plan with covered text segments and missing-glyph diagnostics.</returns>
    public PdfTextFallbackPlan PlanText(string text, string source = "", PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) =>
        PdfTextDiagnostics.PlanEmbeddedFontFallbackText(text, _candidates, source, shapingMode);

    /// <summary>
    /// Finds advanced-layout warnings for text planned against this fallback set, including selected-font OpenType feature gaps.
    /// </summary>
    /// <param name="text">Text to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="shapingMode">Text shaping mode to use when checking fallback font coverage.</param>
    /// <returns>Advanced text layout diagnostics in source order, with duplicate warning families collapsed.</returns>
    public IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayout(string text, string source = "", PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) =>
        AnalyzeAdvancedTextLayout(PlanText(text, source, shapingMode), source);

    /// <summary>
    /// Plans text and converts a fully covered plan into styled rich text runs assigned to this set's generated font slots.
    /// </summary>
    /// <param name="text">Text to inspect and convert.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="styleTemplate">Optional run whose styling is copied to each generated text run.</param>
    /// <param name="shapingMode">Text shaping mode to use when checking fallback font coverage.</param>
    /// <returns>Text runs that can be used with rich paragraphs, lists, tables, panels, and canvas text boxes.</returns>
    public IReadOnlyList<TextRun> PlanTextRuns(string text, string source = "", TextRun? styleTemplate = null, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) =>
        UsesNamedFontFamilies
            ? PlanText(text, source, shapingMode).ToNamedTextRuns(_fontFamilyNames, styleTemplate)
            : PlanText(text, source, shapingMode).ToTextRuns(_fontSlots, styleTemplate);

    /// <summary>
    /// Plans text and returns renderable rich text runs only when every non-layout scalar is covered.
    /// </summary>
    /// <param name="text">Text to inspect and convert.</param>
    /// <param name="runs">Renderable rich text runs when the plan is fully covered; otherwise an empty collection.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="styleTemplate">Optional run whose styling is copied to each generated text run.</param>
    /// <param name="report">Optional conversion report that receives missing-glyph diagnostics from incomplete plans.</param>
    /// <param name="converter">Converter or adapter name used when writing diagnostics to <paramref name="report"/>.</param>
    /// <param name="shapingMode">Text shaping mode to use when checking fallback font coverage.</param>
    /// <returns><c>true</c> when <paramref name="runs"/> contains a fully covered renderable plan; otherwise <c>false</c>.</returns>
    public bool TryPlanTextRuns(
        string text,
        out IReadOnlyList<TextRun> runs,
        string source = "",
        TextRun? styleTemplate = null,
        PdfConversionReport? report = null,
        string converter = "OfficeIMO.Pdf",
        PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) {
        PdfTextFallbackPlan plan = PlanText(text, source, shapingMode);
        report?.AddTextFallbackPlanDiagnostics(plan, converter);
        if (report != null) {
            report.AddTextShapingDiagnostics(AnalyzeAdvancedTextLayout(plan, source), converter);
        }

        if (!plan.IsFullyCovered) {
            runs = Array.Empty<TextRun>();
            return false;
        }

        runs = UsesNamedFontFamilies
            ? plan.ToNamedTextRuns(_fontFamilyNames, styleTemplate)
            : plan.ToTextRuns(_fontSlots, styleTemplate);
        return true;
    }

    private void ValidateCandidates() {
        if (_candidates.Count == 0) {
            throw new ArgumentException("At least one embedded font fallback candidate is required.", "candidates");
        }
    }

    private List<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayout(PdfTextFallbackPlan plan, string source) {
        var diagnostics = new List<PdfTextShapingDiagnostic>();
        var reported = new HashSet<string>(StringComparer.Ordinal);

        AddDiagnostics(
            diagnostics,
            reported,
            PdfTextDiagnostics.AnalyzeAdvancedTextLayout(plan.OriginalText, source));

        foreach (PdfTextFallbackSegment segment in plan.Segments) {
            if (segment.FontIndex < 0 || segment.FontIndex >= _candidates.Count) {
                continue;
            }

            PdfEmbeddedFontFallbackCandidate candidate = _candidates[segment.FontIndex];
            AddDiagnostics(
                diagnostics,
                reported,
                PdfTextDiagnostics.AnalyzeAdvancedTextLayout(
                    segment.Text,
                    candidate.DataSnapshot,
                    source,
                    candidate.FontName,
                    segment.StartIndex));
        }

        return diagnostics;
    }

    private static void AddDiagnostics(List<PdfTextShapingDiagnostic> target, HashSet<string> reported, IReadOnlyList<PdfTextShapingDiagnostic> diagnostics) {
        foreach (PdfTextShapingDiagnostic diagnostic in diagnostics) {
            string key = diagnostic.Code + "|" + diagnostic.Source + "|" + diagnostic.Script;
            if (reported.Add(key)) {
                target.Add(diagnostic);
            }
        }
    }

    private static PdfEmbeddedFontFallbackCandidate CloneCandidate(PdfEmbeddedFontFallbackCandidate candidate) {
        if (candidate == null) {
            throw new ArgumentException("Embedded font fallback candidates cannot contain null entries.", nameof(candidate));
        }

        return candidate;
    }

    private static PdfStandardFont NormalizeFontSlot(PdfStandardFont font) {
        Guard.StandardFont(font, nameof(font), "Fallback font slots must be supported generated standard PDF fonts.");
        return PdfStandardFontMapper.GetFontFamily(font);
    }
}
