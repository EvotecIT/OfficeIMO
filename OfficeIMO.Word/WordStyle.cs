using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Identifies the kind of content a Word style formats.</summary>
public enum WordStyleType {
    /// <summary>Formats whole paragraphs, including their text.</summary>
    Paragraph,
    /// <summary>Formats runs of text inside a paragraph.</summary>
    Character,
    /// <summary>Formats tables, their regions and their text.</summary>
    Table,
    /// <summary>Names a list definition.</summary>
    Numbering
}

/// <summary>Selects which theme font a style uses.</summary>
public enum WordThemeFont {
    /// <summary>The theme's major (headings) font.</summary>
    Major,
    /// <summary>The theme's minor (body) font.</summary>
    Minor
}

/// <summary>
/// A style defined in a <see cref="WordDocument"/>. Reading and setting its properties works directly on the
/// document's style definition, so every paragraph, run or table that uses the style follows the change.
/// </summary>
/// <remarks>
/// Formatting properties return the value declared by this style, not the effective value inherited from its
/// base style or the document defaults. A <see langword="null"/> value inherits; assigning <see langword="null"/>
/// removes the declaration.
/// </remarks>
public sealed class WordStyle {
    private readonly WordDocument _document;

    internal WordStyle(WordDocument document, Style element) {
        _document = document;
        Element = element;
    }

    internal Style Element { get; }

    internal WordDocument Document => _document;

    /// <summary>Gets the identifier that paragraphs, runs and tables store to use this style.</summary>
    public string StyleId => Element.StyleId?.Value ?? string.Empty;

    /// <summary>Gets the kind of content this style formats.</summary>
    public WordStyleType Type => ReadType(Element);

    /// <summary>Gets or sets the style name shown in Word. Names are unique within a document, ignoring case.</summary>
    public string Name {
        get => Element.StyleName?.Val?.Value ?? StyleId;
        set {
            ArgumentException.ThrowIfNullOrWhiteSpace(value);
            EnsureEditable();
            WordStyle? existing = _document.Styles.FindByName(value);
            if (existing != null && !ReferenceEquals(existing.Element, Element))
                throw new ArgumentException($"The document already has a style named '{value}'.", nameof(value));
            Element.StyleName = new StyleName { Val = value };
        }
    }

    /// <summary>Gets or sets the identifier of the style this style inherits from.</summary>
    /// <exception cref="ArgumentException">The base style would make this style inherit from itself.</exception>
    public string? BasedOnStyleId {
        get => Element.BasedOn?.Val?.Value;
        set {
            EnsureEditable();
            if (string.IsNullOrWhiteSpace(value)) {
                Element.BasedOn = null;
                return;
            }
            // A loaded document can already contain a cycle elsewhere in the chain, so stop at the first repeat.
            var visited = new HashSet<string>(StringComparer.Ordinal);
            for (string? current = value; current != null && visited.Add(current); current = _document.Styles.FindById(current)?.BasedOnStyleId) {
                if (string.Equals(current, StyleId, StringComparison.Ordinal))
                    throw new ArgumentException($"Basing '{StyleId}' on '{value}' would make it inherit from itself.", nameof(value));
            }
            Element.BasedOn = new BasedOn { Val = value };
        }
    }

    /// <summary>Gets or sets the identifier of the paragraph style Word uses for the next paragraph.</summary>
    public string? NextStyleId {
        get => Element.NextParagraphStyle?.Val?.Value;
        set {
            EnsureEditable();
            Element.NextParagraphStyle = string.IsNullOrWhiteSpace(value) ? null : new NextParagraphStyle { Val = value };
        }
    }

    /// <summary>
    /// Gets or sets the identifier of the linked style: the character style paired with a paragraph style, or the
    /// paragraph style paired with a character style.
    /// </summary>
    public string? LinkedStyleId {
        get => Element.LinkedStyle?.Val?.Value;
        set {
            EnsureEditable();
            Element.LinkedStyle = string.IsNullOrWhiteSpace(value) ? null : new LinkedStyle { Val = value };
        }
    }

    /// <summary>
    /// Gets or sets whether this is the default style of its type, used by content that names no style.
    /// Making a style the default clears the flag on the previous default of the same type.
    /// </summary>
    public bool IsDefault {
        get => Element.Default?.Value ?? false;
        set {
            EnsureEditable();
            if (value) {
                foreach (WordStyle other in _document.Styles.Where(style => style.Type == Type && style.IsDefault))
                    other.Element.Default = null;
                Element.Default = true;
            } else {
                Element.Default = null;
            }
        }
    }

    /// <summary>Gets whether the style was defined by a user or application rather than being built into Word.</summary>
    public bool IsCustom => Element.CustomStyle?.Value ?? false;

    /// <summary>Gets or sets whether Word shows the style in the Styles gallery.</summary>
    public bool IsPrimary {
        get => Element.PrimaryStyle != null;
        set {
            EnsureEditable();
            Element.PrimaryStyle = value ? new PrimaryStyle() : null;
        }
    }

    // Run formatting ---------------------------------------------------------------------------------------------

    /// <summary>Gets or sets the font family. Reading a theme font returns the family the document theme names.</summary>
    /// <remarks>Setting a family writes it for every script and replaces a theme font selection.</remarks>
    public string? FontFamily {
        get {
            RunFonts? fonts = ReadRun()?.GetFirstChild<RunFonts>();
            if (fonts == null) return null;
            string? literal = fonts.Ascii?.Value ?? fonts.HighAnsi?.Value;
            if (fonts.AsciiTheme == null && fonts.HighAnsiTheme == null) return literal;
            return ReadThemeFont(fonts) is { } role ? _document.Theme.GetFont(role) ?? literal : literal;
        }
        set {
            OpenXmlCompositeElement run = EnsureRun();
            run.RemoveAllChildren<RunFonts>();
            if (string.IsNullOrWhiteSpace(value)) { PruneRun(); return; }
            run.AddChild(new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value, ComplexScript = value }, true);
        }
    }

    /// <summary>Gets or sets which theme font the style uses, so the text follows changes to the document theme.</summary>
    public WordThemeFont? ThemeFont {
        get => ReadThemeFont(ReadRun()?.GetFirstChild<RunFonts>());
        set {
            OpenXmlCompositeElement run = EnsureRun();
            run.RemoveAllChildren<RunFonts>();
            if (!value.HasValue) { PruneRun(); return; }
            bool major = value.Value == WordThemeFont.Major;
            string? family = _document.Theme.GetFont(value.Value);
            run.AddChild(new RunFonts {
                AsciiTheme = major ? ThemeFontValues.MajorHighAnsi : ThemeFontValues.MinorHighAnsi,
                HighAnsiTheme = major ? ThemeFontValues.MajorHighAnsi : ThemeFontValues.MinorHighAnsi,
                EastAsiaTheme = major ? ThemeFontValues.MajorEastAsia : ThemeFontValues.MinorEastAsia,
                ComplexScriptTheme = major ? ThemeFontValues.MajorBidi : ThemeFontValues.MinorBidi,
                // Literal families are what applications without theme support display.
                Ascii = family,
                HighAnsi = family
            }, true);
        }
    }

    /// <summary>Gets or sets the font size, rounded to Word's half-point precision.</summary>
    public OfficeLength? FontSize {
        get => double.TryParse(ReadRun()?.GetFirstChild<FontSize>()?.Val?.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double halfPoints)
            ? OfficeLength.Points(halfPoints / 2d) : null;
        set {
            OpenXmlCompositeElement run = EnsureRun();
            run.RemoveAllChildren<FontSize>();
            run.RemoveAllChildren<FontSizeComplexScript>();
            if (!value.HasValue) { PruneRun(); return; }
            string size = value.Value.ToHalfPoints().ToString(CultureInfo.InvariantCulture);
            run.AddChild(new FontSize { Val = size }, true);
            run.AddChild(new FontSizeComplexScript { Val = size }, true);
        }
    }

    /// <summary>Gets or sets the text colour. A theme colour reads as its current value in the document theme.</summary>
    /// <remarks>Setting an RGB colour replaces a theme colour selection. Automatic colour reads as <see langword="null"/>.</remarks>
    public OfficeColor? Color {
        get => Helpers.ParseNativeColor(WordRunFormatting.ResolveColorHex(ReadRun()?.GetFirstChild<Color>(), _document));
        set {
            OpenXmlCompositeElement run = EnsureRun();
            run.RemoveAllChildren<Color>();
            if (!value.HasValue) { PruneRun(); return; }
            run.AddChild(new Color { Val = value.Value.ToRgbHex() }, true);
        }
    }

    /// <summary>Gets or sets the theme colour of the text, so it follows changes to the document theme.</summary>
    public WordThemeColor? ThemeColor {
        get => ReadRun()?.GetFirstChild<Color>()?.ThemeColor?.Value.ToOfficeEnum();
        set {
            OpenXmlCompositeElement run = EnsureRun();
            run.RemoveAllChildren<Color>();
            if (!value.HasValue) { PruneRun(); return; }
            var color = new Color { ThemeColor = value.Value.ToOpenXml() };
            color.Val = WordRunFormatting.ResolveColorHex(color, _document) ?? "auto";
            run.AddChild(color, true);
        }
    }

    /// <summary>Gets or sets bold text.</summary>
    public bool? Bold {
        get => ReadOnOff(ReadRun()?.GetFirstChild<Bold>());
        set => SetRunOnOff<Bold, BoldComplexScript>(value);
    }

    /// <summary>Gets or sets italic text.</summary>
    public bool? Italic {
        get => ReadOnOff(ReadRun()?.GetFirstChild<Italic>());
        set => SetRunOnOff<Italic, ItalicComplexScript>(value);
    }

    /// <summary>Gets or sets the underline.</summary>
    public WordUnderlineStyle? Underline {
        get => ReadRun()?.GetFirstChild<Underline>()?.Val?.Value.ToOfficeEnum();
        set {
            OpenXmlCompositeElement run = EnsureRun();
            run.RemoveAllChildren<Underline>();
            if (!value.HasValue) { PruneRun(); return; }
            run.AddChild(new Underline { Val = value.Value.ToOpenXml() }, true);
        }
    }

    /// <summary>Gets or sets single strikethrough.</summary>
    public bool? Strikethrough {
        get => ReadOnOff(ReadRun()?.GetFirstChild<Strike>());
        set => SetRunOnOff<Strike, Strike>(value);
    }

    /// <summary>Gets or sets capitalization: all capitals, small capitals, or explicitly neither.</summary>
    public WordCapsStyle? CapsStyle {
        get {
            OpenXmlCompositeElement? run = ReadRun();
            bool? caps = ReadOnOff(run?.GetFirstChild<Caps>());
            bool? smallCaps = ReadOnOff(run?.GetFirstChild<SmallCaps>());
            if (caps == true) return WordCapsStyle.Caps;
            if (smallCaps == true) return WordCapsStyle.SmallCaps;
            return caps.HasValue || smallCaps.HasValue ? WordCapsStyle.None : null;
        }
        set {
            OpenXmlCompositeElement run = EnsureRun();
            run.RemoveAllChildren<Caps>();
            run.RemoveAllChildren<SmallCaps>();
            if (!value.HasValue) { PruneRun(); return; }
            run.AddChild(value.Value == WordCapsStyle.Caps ? new Caps() : new Caps { Val = false }, true);
            run.AddChild(value.Value == WordCapsStyle.SmallCaps ? new SmallCaps() : new SmallCaps { Val = false }, true);
        }
    }

    /// <summary>Gets or sets the highlight colour.</summary>
    public WordHighlightColor? Highlight {
        get => ReadRun()?.GetFirstChild<Highlight>()?.Val?.Value.ToOfficeEnum();
        set {
            OpenXmlCompositeElement run = EnsureRun();
            run.RemoveAllChildren<Highlight>();
            if (!value.HasValue) { PruneRun(); return; }
            run.AddChild(new Highlight { Val = value.Value.ToOpenXml() }, true);
        }
    }

    // Paragraph formatting ---------------------------------------------------------------------------------------

    /// <summary>Gets or sets paragraph alignment.</summary>
    public WordParagraphAlignment? ParagraphAlignment {
        get => ReadParagraph()?.Justification?.Val?.Value.ToOfficeEnum();
        set {
            StyleParagraphProperties paragraph = EnsureParagraph();
            paragraph.Justification = value.HasValue ? new Justification { Val = value.Value.ToOpenXml() } : null;
            PruneParagraph();
        }
    }

    /// <summary>Gets or sets the indentation before the text (left in left-to-right paragraphs).</summary>
    public OfficeLength? IndentationBefore {
        get => ReadTwips(ReadParagraph()?.Indentation?.Left?.Value);
        set => SetIndentation(indentation => indentation.Left = Twips(value));
    }

    /// <summary>Gets or sets the indentation after the text (right in left-to-right paragraphs).</summary>
    public OfficeLength? IndentationAfter {
        get => ReadTwips(ReadParagraph()?.Indentation?.Right?.Value);
        set => SetIndentation(indentation => indentation.Right = Twips(value));
    }

    /// <summary>Gets or sets the first-line indentation. Setting it removes a hanging indentation.</summary>
    public OfficeLength? IndentationFirstLine {
        get => ReadTwips(ReadParagraph()?.Indentation?.FirstLine?.Value);
        set => SetIndentation(indentation => {
            indentation.FirstLine = Twips(value);
            if (value.HasValue) indentation.Hanging = null;
        });
    }

    /// <summary>Gets or sets the hanging indentation. Setting it removes a first-line indentation.</summary>
    public OfficeLength? IndentationHanging {
        get => ReadTwips(ReadParagraph()?.Indentation?.Hanging?.Value);
        set => SetIndentation(indentation => {
            indentation.Hanging = Twips(value);
            if (value.HasValue) indentation.FirstLine = null;
        });
    }

    /// <summary>Gets or sets the spacing before the paragraph.</summary>
    public OfficeLength? LineSpacingBefore {
        get => ReadTwips(ReadParagraph()?.SpacingBetweenLines?.Before?.Value);
        set => SetSpacing(spacing => spacing.Before = Twips(value));
    }

    /// <summary>Gets or sets the spacing after the paragraph.</summary>
    public OfficeLength? LineSpacingAfter {
        get => ReadTwips(ReadParagraph()?.SpacingBetweenLines?.After?.Value);
        set => SetSpacing(spacing => spacing.After = Twips(value));
    }

    /// <summary>Gets or sets line spacing as a multiple, an exact height or a minimum height.</summary>
    public WordLineSpacing? LineSpacing {
        get => WordLineSpacing.Read(ReadParagraph()?.SpacingBetweenLines);
        set => SetSpacing(spacing => {
            spacing.Line = value?.ToNativeValue().ToString(CultureInfo.InvariantCulture);
            spacing.LineRule = value?.Rule.ToOpenXml();
        });
    }

    /// <summary>Gets or sets whether paragraphs start on a new page.</summary>
    public bool? PageBreakBefore {
        get => ReadOnOff(ReadParagraph()?.PageBreakBefore);
        set => SetParagraphOnOff<PageBreakBefore>(value);
    }

    /// <summary>Gets or sets whether paragraphs stay on the same page as the next paragraph.</summary>
    public bool? KeepWithNext {
        get => ReadOnOff(ReadParagraph()?.KeepNext);
        set => SetParagraphOnOff<KeepNext>(value);
    }

    /// <summary>Gets or sets whether all lines of a paragraph stay on one page.</summary>
    public bool? KeepLinesTogether {
        get => ReadOnOff(ReadParagraph()?.KeepLines);
        set => SetParagraphOnOff<KeepLines>(value);
    }

    /// <summary>Gets or sets widow and orphan control.</summary>
    public bool? AvoidWidowAndOrphan {
        get => ReadOnOff(ReadParagraph()?.WidowControl);
        set => SetParagraphOnOff<WidowControl>(value);
    }

    /// <summary>Gets or sets whether spacing between paragraphs of this style is suppressed.</summary>
    public bool? ContextualSpacing {
        get => ReadOnOff(ReadParagraph()?.ContextualSpacing);
        set => SetParagraphOnOff<ContextualSpacing>(value);
    }

    /// <summary>Gets or sets whether line numbers are suppressed for paragraphs of this style.</summary>
    public bool? SuppressLineNumbers {
        get => ReadOnOff(ReadParagraph()?.SuppressLineNumbers);
        set => SetParagraphOnOff<SuppressLineNumbers>(value);
    }

    /// <summary>Gets or sets whether automatic hyphenation is suppressed for paragraphs of this style.</summary>
    public bool? SuppressAutoHyphens {
        get => ReadOnOff(ReadParagraph()?.SuppressAutoHyphens);
        set => SetParagraphOnOff<SuppressAutoHyphens>(value);
    }

    /// <summary>Gets or sets whether the indents before and after the text are used as inside and outside indents.</summary>
    public bool? MirrorIndents {
        get => ReadOnOff(ReadParagraph()?.MirrorIndents);
        set => SetParagraphOnOff<MirrorIndents>(value);
    }

    /// <summary>Gets or sets the outline level: 0-8 for heading levels 1-9, or 9 for body text.</summary>
    public int? OutlineLevel {
        get => ReadParagraph()?.OutlineLevel?.Val?.Value;
        set {
            if (value is < 0 or > 9) throw new ArgumentOutOfRangeException(nameof(value), "Outline level must be between 0 and 9.");
            StyleParagraphProperties paragraph = EnsureParagraph();
            paragraph.OutlineLevel = value.HasValue ? new OutlineLevel { Val = value.Value } : null;
            PruneParagraph();
        }
    }

    /// <summary>Gets or sets the paragraph background colour.</summary>
    public OfficeColor? ShadingFillColor {
        get => Helpers.ParseNativeColor(WordRunFormatting.ResolveShadingFillColorHex(ReadParagraph()?.Shading, _document));
        set {
            StyleParagraphProperties paragraph = EnsureParagraph();
            paragraph.Shading = value.HasValue
                ? new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = value.Value.ToRgbHex() }
                : null;
            PruneParagraph();
        }
    }

    /// <summary>Gets the paragraph borders declared by this style.</summary>
    /// <exception cref="InvalidOperationException">The style is a character or numbering style.</exception>
    public WordParagraphBorders Borders {
        get {
            RequireParagraphFormatting();
            return new WordParagraphBorders(
                () => Element.StyleParagraphProperties,
                EnsureParagraph);
        }
    }

    /// <summary>Gets the list that numbers paragraphs using this style, or <see langword="null"/>.</summary>
    public WordList? List => ReadParagraph()?.NumberingProperties?.NumberingId?.Val?.Value is int id && id > 0
        ? new WordList(_document, id) : null;

    /// <summary>Gets the list level that numbers paragraphs using this style, or <see langword="null"/>.</summary>
    public int? ListLevel {
        get {
            NumberingProperties? numbering = ReadParagraph()?.NumberingProperties;
            if (numbering?.NumberingId == null) return null;
            return numbering.NumberingLevelReference?.Val?.Value ?? 0;
        }
    }

    /// <summary>Numbers paragraphs that use this style with a level of a list, or removes the numbering.</summary>
    /// <param name="list">The list to number with, or <see langword="null"/> to remove numbering.</param>
    /// <param name="level">The zero-based list level.</param>
    /// <returns>This style.</returns>
    public WordStyle SetList(WordList? list, int level = 0) {
        if (level is < 0 or > 8) throw new ArgumentOutOfRangeException(nameof(level), "List level must be between 0 and 8.");
        StyleParagraphProperties paragraph = EnsureParagraph();
        if (list == null) {
            paragraph.NumberingProperties = null;
        } else {
            var numbering = new NumberingProperties();
            if (level > 0) numbering.NumberingLevelReference = new NumberingLevelReference { Val = level };
            numbering.NumberingId = new NumberingId { Val = list.NumberId };
            paragraph.NumberingProperties = numbering;
        }
        PruneParagraph();
        return this;
    }

    // Table formatting -------------------------------------------------------------------------------------------

    /// <summary>Gets the formatting this table style applies to one region of a table.</summary>
    /// <param name="region">The table region.</param>
    /// <exception cref="InvalidOperationException">The style is not a table style.</exception>
    public WordTableStyleRegionFormat GetTableRegion(WordTableStyleRegion region) {
        if (Type != WordStyleType.Table) throw new InvalidOperationException($"'{StyleId}' is not a table style.");
        return new WordTableStyleRegionFormat(this, region);
    }

    /// <summary>Sets the outside and inside borders that this table style draws around and between cells.</summary>
    /// <param name="style">The border line style; <see cref="WordBorderStyle.None"/> removes the borders.</param>
    /// <param name="size">The border width, rounded to an eighth of a point.</param>
    /// <param name="color">The border colour.</param>
    /// <returns>This style.</returns>
    public WordStyle SetTableBorders(WordBorderStyle style, OfficeLength size, OfficeColor color) {
        if (Type != WordStyleType.Table) throw new InvalidOperationException($"'{StyleId}' is not a table style.");
        EnsureEditable();
        uint width = WordBorderLengths.ToSize(size) ?? 0;
        StyleTableProperties table = Element.StyleTableProperties ??= new StyleTableProperties();
        table.RemoveAllChildren<TableBorders>();
        if (style != WordBorderStyle.None) {
            BorderValues value = style.ToOpenXml();
            string hex = color.ToRgbHex();
            T Border<T>() where T : BorderType, new() => new() { Val = value, Size = width, Space = 0U, Color = hex };
            table.AddChild(new TableBorders(Border<TopBorder>(), Border<LeftBorder>(), Border<BottomBorder>(),
                Border<RightBorder>(), Border<InsideHorizontalBorder>(), Border<InsideVerticalBorder>()), true);
        }
        return this;
    }

    // Shared plumbing --------------------------------------------------------------------------------------------

    internal static WordStyleType ReadType(Style style) {
        StyleValues? type = style.Type?.Value;
        if (type == StyleValues.Character) return WordStyleType.Character;
        if (type == StyleValues.Table) return WordStyleType.Table;
        if (type == StyleValues.Numbering) return WordStyleType.Numbering;
        return WordStyleType.Paragraph;
    }

    internal void EnsureEditable() {
        if (_document.FileOpenAccess == FileAccess.Read) throw new InvalidOperationException("A read-only document's styles cannot be changed.");
        if (Element.Parent == null) throw new InvalidOperationException($"The style '{StyleId}' was removed from the document.");
    }

    private OpenXmlCompositeElement? ReadRun() => Element.StyleRunProperties;

    private OpenXmlCompositeElement EnsureRun() {
        if (Type == WordStyleType.Numbering) throw new InvalidOperationException($"Numbering style '{StyleId}' has no text formatting.");
        EnsureEditable();
        return Element.StyleRunProperties ??= new StyleRunProperties();
    }

    private void PruneRun() {
        if (Element.StyleRunProperties is { HasChildren: false } run) run.Remove();
    }

    private void SetRunOnOff<T, TComplex>(bool? value) where T : OnOffType, new() where TComplex : OnOffType, new() {
        OpenXmlCompositeElement run = EnsureRun();
        run.RemoveAllChildren<T>();
        run.RemoveAllChildren<TComplex>();
        if (!value.HasValue) { PruneRun(); return; }
        run.AddChild(value.Value ? new T() : new T { Val = false }, true);
        if (typeof(TComplex) != typeof(T)) run.AddChild(value.Value ? new TComplex() : new TComplex { Val = false }, true);
    }

    private StyleParagraphProperties? ReadParagraph() => Element.StyleParagraphProperties;

    private void RequireParagraphFormatting() {
        if (Type is WordStyleType.Character or WordStyleType.Numbering)
            throw new InvalidOperationException($"'{StyleId}' is a {Type.ToString().ToLowerInvariant()} style and has no paragraph formatting.");
    }

    private StyleParagraphProperties EnsureParagraph() {
        RequireParagraphFormatting();
        EnsureEditable();
        return Element.StyleParagraphProperties ??= new StyleParagraphProperties();
    }

    private void PruneParagraph() {
        if (Element.StyleParagraphProperties is { } paragraph) {
            if (paragraph.Indentation is { HasAttributes: false }) paragraph.Indentation = null;
            if (paragraph.SpacingBetweenLines is { HasAttributes: false }) paragraph.SpacingBetweenLines = null;
            if (!paragraph.HasChildren) paragraph.Remove();
        }
    }

    private void SetParagraphOnOff<T>(bool? value) where T : OnOffType, new() {
        StyleParagraphProperties paragraph = EnsureParagraph();
        paragraph.RemoveAllChildren<T>();
        if (value.HasValue) paragraph.AddChild(value.Value ? new T() : new T { Val = false }, true);
        PruneParagraph();
    }

    private void SetIndentation(Action<Indentation> change) {
        StyleParagraphProperties paragraph = EnsureParagraph();
        change(paragraph.Indentation ??= new Indentation());
        PruneParagraph();
    }

    private void SetSpacing(Action<SpacingBetweenLines> change) {
        StyleParagraphProperties paragraph = EnsureParagraph();
        change(paragraph.SpacingBetweenLines ??= new SpacingBetweenLines());
        PruneParagraph();
    }

    private static WordThemeFont? ReadThemeFont(RunFonts? fonts) {
        string? token = fonts?.AsciiTheme?.InnerText ?? fonts?.HighAnsiTheme?.InnerText;
        if (token == null) return null;
        if (token.StartsWith("major", StringComparison.Ordinal)) return WordThemeFont.Major;
        if (token.StartsWith("minor", StringComparison.Ordinal)) return WordThemeFont.Minor;
        return null;
    }

    internal static bool? ReadOnOff(OnOffType? value) => value == null ? null : value.Val?.Value ?? true;

    private static OfficeLength? ReadTwips(string? value) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips)
            ? OfficeLength.Emus((long)twips * 635) : null;

    private static string? Twips(OfficeLength? value) =>
        value.HasValue ? checked((int)Math.Round((double)value.Value.ToTwips(), MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture) : null;

    /// <inheritdoc />
    public override string ToString() => StyleId;
}
