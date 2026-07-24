namespace OfficeIMO.Pdf;

/// <summary>
/// Describes how text can be split across embedded fonts before generated PDF rendering.
/// </summary>
public sealed class PdfTextFallbackPlan {
    internal PdfTextFallbackPlan(string originalText, IReadOnlyList<PdfTextFallbackSegment> segments, IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics) {
        OriginalText = originalText ?? string.Empty;
        Segments = segments ?? throw new ArgumentNullException(nameof(segments));
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
    }

    /// <summary>Original text used to create this fallback plan.</summary>
    public string OriginalText { get; }

    /// <summary>Contiguous text segments that can be rendered by one embedded font candidate.</summary>
    public IReadOnlyList<PdfTextFallbackSegment> Segments { get; }

    /// <summary>Characters not covered by any candidate font, in source order.</summary>
    public IReadOnlyList<PdfTextEncodingDiagnostic> Diagnostics { get; }

    /// <summary>True when every non-layout text scalar is covered by one of the candidate fonts.</summary>
    public bool IsFullyCovered => Diagnostics.Count == 0;

    /// <summary>
    /// Converts a fully covered fallback plan into rich text runs assigned to generated PDF font slots.
    /// </summary>
    /// <param name="fontSlots">Generated standard-font slots ordered the same way as the fallback candidates.</param>
    /// <param name="styleTemplate">Optional run whose styling is copied to each generated text run.</param>
    /// <returns>Text runs that can be used with rich paragraphs, lists, tables, panels, and canvas text boxes.</returns>
    public IReadOnlyList<TextRun> ToTextRuns(IReadOnlyList<PdfStandardFont> fontSlots, TextRun? styleTemplate = null) {
        Guard.NotNull(fontSlots, nameof(fontSlots));
        return ToTextRuns(index => {
            if (index < 0 || index >= fontSlots.Count) {
                throw new ArgumentException("Fallback font slot mapping is missing an entry for candidate index " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(fontSlots));
            }

            PdfStandardFont slot = fontSlots[index];
            Guard.StandardFont(slot, nameof(fontSlots), "Fallback font slots must be supported generated standard PDF fonts.");
            return slot;
        }, styleTemplate);
    }

    /// <summary>
    /// Converts a fully covered fallback plan into rich text runs assigned to generated PDF font slots.
    /// </summary>
    /// <param name="fontSlots">Generated standard-font slots keyed by fallback candidate index.</param>
    /// <param name="styleTemplate">Optional run whose styling is copied to each generated text run.</param>
    /// <returns>Text runs that can be used with rich paragraphs, lists, tables, panels, and canvas text boxes.</returns>
    public IReadOnlyList<TextRun> ToTextRuns(IReadOnlyDictionary<int, PdfStandardFont> fontSlots, TextRun? styleTemplate = null) {
        Guard.NotNull(fontSlots, nameof(fontSlots));
        return ToTextRuns(index => {
            if (!fontSlots.TryGetValue(index, out PdfStandardFont slot)) {
                throw new ArgumentException("Fallback font slot mapping is missing an entry for candidate index " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(fontSlots));
            }

            Guard.StandardFont(slot, nameof(fontSlots), "Fallback font slots must be supported generated standard PDF fonts.");
            return slot;
        }, styleTemplate);
    }

    /// <summary>
    /// Converts a fully covered fallback plan into rich text runs assigned to registered named font families.
    /// </summary>
    /// <param name="fontFamilyNames">Registered named font families ordered the same way as the fallback candidates.</param>
    /// <param name="styleTemplate">Optional run whose styling is copied to each generated text run.</param>
    /// <returns>Text runs that can be used with rich paragraphs, lists, tables, panels, and canvas text boxes.</returns>
    public IReadOnlyList<TextRun> ToNamedTextRuns(IReadOnlyList<string> fontFamilyNames, TextRun? styleTemplate = null) {
        Guard.NotNull(fontFamilyNames, nameof(fontFamilyNames));
        return ToNamedTextRuns(index => {
            if (index < 0 || index >= fontFamilyNames.Count) {
                throw new ArgumentException("Fallback named font mapping is missing an entry for candidate index " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(fontFamilyNames));
            }

            string fontFamilyName = fontFamilyNames[index];
            Guard.NotNullOrWhiteSpace(fontFamilyName, nameof(fontFamilyNames));
            return fontFamilyName.Trim();
        }, styleTemplate);
    }

    private System.Collections.ObjectModel.ReadOnlyCollection<TextRun> ToTextRuns(Func<int, PdfStandardFont> resolveFont, TextRun? styleTemplate) {
        if (!IsFullyCovered) {
            throw new InvalidOperationException("Cannot convert an incomplete embedded-font fallback plan to renderable text runs. Inspect Diagnostics and add fallback font coverage first.");
        }

        var runs = new List<TextRun>();
        int cursor = 0;
        foreach (PdfTextFallbackSegment segment in Segments) {
            if (segment.StartIndex > cursor) {
                AddLayoutControlRuns(runs, OriginalText.Substring(cursor, segment.StartIndex - cursor));
            }

            runs.Add(CreateStyledRun(segment.Text, resolveFont(segment.FontIndex), styleTemplate));
            cursor = segment.StartIndex + segment.Length;
        }

        if (cursor < OriginalText.Length) {
            AddLayoutControlRuns(runs, OriginalText.Substring(cursor));
        }

        return runs.AsReadOnly();
    }

    private System.Collections.ObjectModel.ReadOnlyCollection<TextRun> ToNamedTextRuns(Func<int, string> resolveFontFamily, TextRun? styleTemplate) {
        if (!IsFullyCovered) {
            throw new InvalidOperationException("Cannot convert an incomplete embedded-font fallback plan to renderable text runs. Inspect Diagnostics and add fallback font coverage first.");
        }

        var runs = new List<TextRun>();
        int cursor = 0;
        foreach (PdfTextFallbackSegment segment in Segments) {
            if (segment.StartIndex > cursor) {
                AddLayoutControlRuns(runs, OriginalText.Substring(cursor, segment.StartIndex - cursor));
            }

            runs.Add(CreateStyledNamedRun(segment.Text, resolveFontFamily(segment.FontIndex), styleTemplate));
            cursor = segment.StartIndex + segment.Length;
        }

        if (cursor < OriginalText.Length) {
            AddLayoutControlRuns(runs, OriginalText.Substring(cursor));
        }

        return runs.AsReadOnly();
    }

    private static TextRun CreateStyledRun(string text, PdfStandardFont font, TextRun? styleTemplate) {
        if (styleTemplate == null) {
            return TextRun.Normal(text, font: font);
        }

        bool keepLink = !string.IsNullOrWhiteSpace(text) &&
            (styleTemplate.LinkUri != null || styleTemplate.LinkDestinationName != null);

        return new TextRun(
            text,
            styleTemplate.Bold,
            styleTemplate.Underline,
            styleTemplate.Color,
            styleTemplate.Italic,
            styleTemplate.Strike,
            styleTemplate.FontSize,
            font,
            keepLink ? styleTemplate.LinkUri : null,
            keepLink ? styleTemplate.LinkContents : null,
            styleTemplate.Baseline,
            keepLink ? styleTemplate.LinkDestinationName : null,
            backgroundColor: styleTemplate.BackgroundColor);
    }

    private static TextRun CreateStyledNamedRun(string text, string fontFamily, TextRun? styleTemplate) {
        if (styleTemplate == null) {
            return TextRun.Normal(text, fontFamily: fontFamily);
        }

        bool keepLink = !string.IsNullOrWhiteSpace(text) &&
            (styleTemplate.LinkUri != null || styleTemplate.LinkDestinationName != null);

        return new TextRun(
            text,
            styleTemplate.Bold,
            styleTemplate.Underline,
            styleTemplate.Color,
            styleTemplate.Italic,
            styleTemplate.Strike,
            styleTemplate.FontSize,
            styleTemplate.Font,
            keepLink ? styleTemplate.LinkUri : null,
            keepLink ? styleTemplate.LinkContents : null,
            styleTemplate.Baseline,
            keepLink ? styleTemplate.LinkDestinationName : null,
            backgroundColor: styleTemplate.BackgroundColor,
            fontFamily: fontFamily);
    }

    private static void AddLayoutControlRuns(List<TextRun> runs, string text) {
        for (int i = 0; i < text.Length; i++) {
            char ch = text[i];
            if (ch == '\n' || ch == '\r') {
                runs.Add(TextRun.LineBreak());
                if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') {
                    i++;
                }
            } else if (ch == '\t') {
                runs.Add(TextRun.Tab());
            } else if (!char.IsControl(ch)) {
                runs.Add(TextRun.Normal(ch.ToString()));
            }
        }
    }
}

/// <summary>
/// Describes one contiguous text segment covered by one embedded TrueType font candidate.
/// </summary>
public sealed class PdfTextFallbackSegment {
    internal PdfTextFallbackSegment(string text, int startIndex, int length, int fontIndex, string fontName) {
        Text = text ?? string.Empty;
        StartIndex = startIndex;
        Length = length;
        FontIndex = fontIndex;
        FontName = fontName ?? string.Empty;
    }

    /// <summary>Segment text.</summary>
    public string Text { get; }

    /// <summary>UTF-16 start index in the original text.</summary>
    public int StartIndex { get; }

    /// <summary>UTF-16 length in the original text.</summary>
    public int Length { get; }

    /// <summary>Zero-based index of the selected fallback candidate.</summary>
    public int FontIndex { get; }

    /// <summary>Display name of the selected fallback candidate.</summary>
    public string FontName { get; }
}
