using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word;

/// <summary>
/// Describes a paragraph style without exposing Open XML SDK types.
/// </summary>
public sealed class WordParagraphStyleDefinition {
    private readonly Style? _template;

    /// <summary>Creates a paragraph style definition.</summary>
    /// <param name="styleId">Identifier stored on paragraphs that use the style.</param>
    public WordParagraphStyleDefinition(string styleId) {
        if (string.IsNullOrWhiteSpace(styleId)) {
            throw new ArgumentException("Style ID cannot be empty.", nameof(styleId));
        }

        StyleId = styleId;
        Name = styleId;
    }

    internal WordParagraphStyleDefinition(Style style) {
        _template = (Style)style.CloneNode(true);
        StyleId = style.StyleId?.Value ?? throw new InvalidOperationException("The paragraph style does not have an ID.");
        Name = style.StyleName?.Val?.Value ?? StyleId;
        BasedOnStyleId = style.BasedOn?.Val?.Value;
        NextStyleId = style.NextParagraphStyle?.Val?.Value;
        IsDefault = style.Default?.Value ?? false;
        IsPrimary = style.PrimaryStyle != null;

        StyleParagraphProperties? paragraph = style.StyleParagraphProperties;
        Alignment = paragraph?.Justification?.Val?.Value.ToOfficeEnum();
        LeftIndentTwips = ReadInt32(paragraph?.Indentation?.Left?.Value);
        SpacingBeforeTwips = ReadInt32(paragraph?.SpacingBetweenLines?.Before?.Value);
        SpacingAfterTwips = ReadInt32(paragraph?.SpacingBetweenLines?.After?.Value);
        PageBreakBefore = ReadOnOff(paragraph?.PageBreakBefore);
        KeepWithNext = ReadOnOff(paragraph?.KeepNext);
        KeepLinesTogether = ReadOnOff(paragraph?.KeepLines);
        AvoidWidowAndOrphan = ReadOnOff(paragraph?.WidowControl);
        ContextualSpacing = ReadOnOff(paragraph?.ContextualSpacing);
        SuppressLineNumbers = ReadOnOff(paragraph?.SuppressLineNumbers);
        SuppressAutoHyphens = ReadOnOff(paragraph?.SuppressAutoHyphens);
        MirrorIndents = ReadOnOff(paragraph?.MirrorIndents);
        OutlineLevel = paragraph?.OutlineLevel?.Val?.Value;

        StyleRunProperties? run = style.StyleRunProperties;
        FontName = run?.RunFonts?.Ascii?.Value ?? run?.RunFonts?.HighAnsi?.Value;
        ColorHex = run?.Color?.Val?.Value;
        FontSizePoints = ReadHalfPoints(run?.FontSize?.Val?.Value);
        Bold = ReadOnOff(run?.Bold);
        Italic = ReadOnOff(run?.Italic);
    }

    /// <summary>Gets the identifier stored on paragraphs that use the style.</summary>
    public string StyleId { get; }

    /// <summary>Gets or sets the friendly style name.</summary>
    public string Name { get; set; }

    /// <summary>Gets or sets the base paragraph style identifier.</summary>
    public string? BasedOnStyleId { get; set; }

    /// <summary>Gets or sets the style identifier used by the following paragraph.</summary>
    public string? NextStyleId { get; set; }

    /// <summary>Gets or sets the font family.</summary>
    public string? FontName { get; set; }

    /// <summary>Gets or sets the font size in points.</summary>
    public double? FontSizePoints { get; set; }

    /// <summary>Gets or sets the six-digit RGB text color.</summary>
    public string? ColorHex { get; set; }

    /// <summary>Gets or sets whether text is bold.</summary>
    public bool? Bold { get; set; }

    /// <summary>Gets or sets whether text is italic.</summary>
    public bool? Italic { get; set; }

    /// <summary>Gets or sets paragraph alignment.</summary>
    public WordParagraphAlignment? Alignment { get; set; }

    /// <summary>Gets or sets the left indentation in twentieths of a point.</summary>
    public int? LeftIndentTwips { get; set; }

    /// <summary>Gets or sets spacing before the paragraph in twentieths of a point.</summary>
    public int? SpacingBeforeTwips { get; set; }

    /// <summary>Gets or sets spacing after the paragraph in twentieths of a point.</summary>
    public int? SpacingAfterTwips { get; set; }

    /// <summary>Gets or sets starting paragraphs on a new page; null inherits from the base style.</summary>
    public bool? PageBreakBefore { get; set; }

    /// <summary>Gets or sets keeping paragraphs with their following paragraphs; null inherits from the base style.</summary>
    public bool? KeepWithNext { get; set; }

    /// <summary>Gets or sets keeping all paragraph lines together; null inherits from the base style.</summary>
    public bool? KeepLinesTogether { get; set; }

    /// <summary>Gets or sets widow/orphan control; null inherits from the base style.</summary>
    public bool? AvoidWidowAndOrphan { get; set; }

    /// <summary>Gets or sets suppression of spacing between paragraphs of the same style; null inherits.</summary>
    public bool? ContextualSpacing { get; set; }

    /// <summary>Gets or sets paragraph line-number suppression; null inherits.</summary>
    public bool? SuppressLineNumbers { get; set; }

    /// <summary>Gets or sets paragraph automatic-hyphenation suppression; null inherits.</summary>
    public bool? SuppressAutoHyphens { get; set; }

    /// <summary>Gets or sets use of left/right indents as inside/outside indents; null inherits.</summary>
    public bool? MirrorIndents { get; set; }

    private int? _outlineLevel;

    /// <summary>Gets or sets outline level: 0-8 for heading levels 1-9, 9 for body text, or null to inherit.</summary>
    public int? OutlineLevel {
        get => _outlineLevel;
        set {
            if (value is < 0 or > 9) throw new ArgumentOutOfRangeException(nameof(value), "Outline level must be between 0 and 9.");
            _outlineLevel = value;
        }
    }

    /// <summary>Gets or sets whether this is the default paragraph style.</summary>
    public bool IsDefault { get; set; }

    /// <summary>Gets or sets whether Word presents this style as a primary style.</summary>
    public bool IsPrimary { get; set; }

    internal Style ToOpenXml() {
        Style style = _template == null
            ? new Style { Type = StyleValues.Paragraph }
            : (Style)_template.CloneNode(true);

        style.Type = StyleValues.Paragraph;
        style.StyleId = StyleId;
        style.Default = IsDefault;
        SetChild(style, string.IsNullOrWhiteSpace(Name) ? null : new StyleName { Val = Name });
        SetChild(style, string.IsNullOrWhiteSpace(BasedOnStyleId) ? null : new BasedOn { Val = BasedOnStyleId });
        SetChild(style, string.IsNullOrWhiteSpace(NextStyleId) ? null : new NextParagraphStyle { Val = NextStyleId });
        SetChild(style, IsPrimary ? new PrimaryStyle() : null);

        bool hasParagraphFormatting = Alignment.HasValue || LeftIndentTwips.HasValue ||
            SpacingBeforeTwips.HasValue || SpacingAfterTwips.HasValue;
        StyleParagraphProperties? paragraph = style.StyleParagraphProperties;
        if (hasParagraphFormatting) {
            paragraph ??= style.StyleParagraphProperties = new StyleParagraphProperties();
            SetChild(paragraph, Alignment.HasValue ? new Justification { Val = Alignment.Value.ToOpenXml() } : null);
            if (LeftIndentTwips.HasValue) {
                paragraph.Indentation ??= new Indentation();
                paragraph.Indentation.Left = LeftIndentTwips.Value.ToString(CultureInfo.InvariantCulture);
            }
            if (SpacingBeforeTwips.HasValue || SpacingAfterTwips.HasValue) {
                paragraph.SpacingBetweenLines ??= new SpacingBetweenLines();
                if (SpacingBeforeTwips.HasValue) {
                    paragraph.SpacingBetweenLines.Before = SpacingBeforeTwips.Value.ToString(CultureInfo.InvariantCulture);
                }
                if (SpacingAfterTwips.HasValue) {
                    paragraph.SpacingBetweenLines.After = SpacingAfterTwips.Value.ToString(CultureInfo.InvariantCulture);
                }
            }
        }

        bool hasPaginationFormatting = PageBreakBefore.HasValue || KeepWithNext.HasValue ||
            KeepLinesTogether.HasValue || AvoidWidowAndOrphan.HasValue || ContextualSpacing.HasValue ||
            SuppressLineNumbers.HasValue || SuppressAutoHyphens.HasValue || MirrorIndents.HasValue || OutlineLevel.HasValue;
        if (paragraph != null || hasPaginationFormatting) {
            paragraph ??= style.StyleParagraphProperties = new StyleParagraphProperties();
            SetParagraphOnOff<PageBreakBefore>(paragraph, PageBreakBefore);
            SetParagraphOnOff<KeepNext>(paragraph, KeepWithNext);
            SetParagraphOnOff<KeepLines>(paragraph, KeepLinesTogether);
            SetParagraphOnOff<WidowControl>(paragraph, AvoidWidowAndOrphan);
            SetParagraphOnOff<ContextualSpacing>(paragraph, ContextualSpacing);
            SetParagraphOnOff<SuppressLineNumbers>(paragraph, SuppressLineNumbers);
            SetParagraphOnOff<SuppressAutoHyphens>(paragraph, SuppressAutoHyphens);
            SetParagraphOnOff<MirrorIndents>(paragraph, MirrorIndents);
            paragraph.OutlineLevel = OutlineLevel.HasValue ? new OutlineLevel { Val = OutlineLevel.Value } : null;
        }

        bool hasRunFormatting = !string.IsNullOrWhiteSpace(FontName) || FontSizePoints.HasValue ||
            !string.IsNullOrWhiteSpace(ColorHex) || Bold.HasValue || Italic.HasValue;
        StyleRunProperties? run = style.StyleRunProperties;
        if (hasRunFormatting) {
            run ??= style.StyleRunProperties = new StyleRunProperties();
            if (!string.IsNullOrWhiteSpace(FontName)) {
                run.RunFonts = new RunFonts {
                    Ascii = FontName,
                    HighAnsi = FontName,
                    ComplexScript = FontName,
                    EastAsia = FontName
                };
            }
            if (FontSizePoints.HasValue) {
                string halfPoints = Math.Round(FontSizePoints.Value * 2, MidpointRounding.AwayFromZero)
                    .ToString(CultureInfo.InvariantCulture);
                run.FontSize = new FontSize { Val = halfPoints };
                run.FontSizeComplexScript = new FontSizeComplexScript { Val = halfPoints };
            }
            if (!string.IsNullOrWhiteSpace(ColorHex)) {
                run.Color = new Color { Val = ColorHex!.Trim().TrimStart('#') };
            }
            SetChild(run, Bold.HasValue ? new Bold { Val = Bold.Value } : null);
            SetChild(run, Italic.HasValue ? new Italic { Val = Italic.Value } : null);
        }

        return style;
    }

    private static void SetChild<T>(OpenXmlCompositeElement parent, T? child) where T : OpenXmlElement {
        parent.RemoveAllChildren<T>();
        if (child != null) {
            parent.AddChild(child, true);
        }
    }

    private static int? ReadInt32(string? value) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result) ? result : null;

    private static double? ReadHalfPoints(string? value) =>
        double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out double result) ? result / 2d : null;

    private static bool? ReadOnOff(OnOffType? value) => value == null ? null : value.Val?.Value ?? true;

    private static void SetParagraphOnOff<T>(StyleParagraphProperties paragraph, bool? value) where T : OnOffType, new() {
        paragraph.RemoveAllChildren<T>();
        if (value.HasValue) paragraph.AddChild(new T { Val = value.Value }, true);
    }
}
