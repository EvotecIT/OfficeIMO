using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// The document theme: its colour scheme and its major (headings) and minor (body) fonts. Text, borders and
/// shading that select a theme colour or theme font follow changes made here.
/// </summary>
/// <remarks>
/// Changing the theme also refreshes the literal colour and font values Word stores next to each theme
/// selection, so applications that ignore the theme show the same result.
/// </remarks>
public sealed class WordTheme {
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private readonly WordDocument _document;

    internal WordTheme(WordDocument document) {
        _document = document;
    }

    private A.Theme? Native => _document._wordprocessingDocument?.MainDocumentPart?.ThemePart?.Theme;

    /// <summary>Gets or sets the theme name.</summary>
    public string? Name {
        get => Native?.Name?.Value;
        set {
            A.Theme theme = EnsureTheme();
            theme.Name = value;
        }
    }

    /// <summary>Gets a colour of the colour scheme, or <see langword="null"/> when the document has no theme.</summary>
    /// <param name="color">The scheme slot. Text and background aliases read the dark and light slots.</param>
    public OfficeColor? GetColor(WordThemeColor color) =>
        OfficeIMO.OpenXml.Internal.OfficeOpenXmlThemeColorResolver.ResolveSchemeColor(Native?.ThemeElements?.ColorScheme, SlotToken(color));

    /// <summary>Sets a colour of the colour scheme.</summary>
    /// <param name="color">The scheme slot. Text and background aliases set the dark and light slots.</param>
    /// <param name="value">The new colour.</param>
    /// <returns>This theme.</returns>
    public WordTheme SetColor(WordThemeColor color, OfficeColor value) {
        string token = SlotToken(color);
        A.ColorScheme scheme = EnsureTheme().ThemeElements!.ColorScheme ??= new A.ColorScheme { Name = "Custom" };
        OpenXmlCompositeElement slot = token switch {
            "dk1" => scheme.Dark1Color ??= new A.Dark1Color(),
            "lt1" => scheme.Light1Color ??= new A.Light1Color(),
            "dk2" => scheme.Dark2Color ??= new A.Dark2Color(),
            "lt2" => scheme.Light2Color ??= new A.Light2Color(),
            "accent1" => scheme.Accent1Color ??= new A.Accent1Color(),
            "accent2" => scheme.Accent2Color ??= new A.Accent2Color(),
            "accent3" => scheme.Accent3Color ??= new A.Accent3Color(),
            "accent4" => scheme.Accent4Color ??= new A.Accent4Color(),
            "accent5" => scheme.Accent5Color ??= new A.Accent5Color(),
            "accent6" => scheme.Accent6Color ??= new A.Accent6Color(),
            "hlink" => scheme.Hyperlink ??= new A.Hyperlink(),
            _ => scheme.FollowedHyperlinkColor ??= new A.FollowedHyperlinkColor()
        };
        slot.RemoveAllChildren();
        slot.AppendChild(new A.RgbColorModelHex { Val = value.ToRgbHex() });
        RefreshFallbacks();
        return this;
    }

    /// <summary>Gets or sets the Latin typeface of the major (headings) font.</summary>
    public string? MajorFont {
        get => GetFont(WordThemeFont.Major);
        set => SetFont(WordThemeFont.Major, value);
    }

    /// <summary>Gets or sets the Latin typeface of the minor (body) font.</summary>
    public string? MinorFont {
        get => GetFont(WordThemeFont.Minor);
        set => SetFont(WordThemeFont.Minor, value);
    }

    internal string? GetFont(WordThemeFont role) {
        A.FontScheme? fonts = Native?.ThemeElements?.FontScheme;
        A.FontCollectionType? collection = role == WordThemeFont.Major ? fonts?.MajorFont : fonts?.MinorFont;
        string? typeface = collection?.LatinFont?.Typeface?.Value;
        return string.IsNullOrEmpty(typeface) ? null : typeface;
    }

    private void SetFont(WordThemeFont role, string? typeface) {
        ArgumentException.ThrowIfNullOrWhiteSpace(typeface);
        A.ThemeElements elements = EnsureTheme().ThemeElements!;
        A.FontScheme fonts = elements.FontScheme ??= new A.FontScheme { Name = "Custom" };
        A.FontCollectionType collection = role == WordThemeFont.Major
            ? (A.FontCollectionType)(fonts.MajorFont ??= new A.MajorFont(new A.LatinFont(), new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }))
            : fonts.MinorFont ??= new A.MinorFont(new A.LatinFont(), new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty });
        A.LatinFont latin = collection.LatinFont ??= new A.LatinFont();
        latin.Typeface = typeface;
        // The panose classification described the previous typeface.
        latin.Panose = null;
        RefreshFallbacks();
    }

    /// <summary>Replaces this document's theme with a copy of another document's theme.</summary>
    /// <param name="source">The document or template whose theme to copy. It is not changed; a read-only load is enough.</param>
    /// <exception cref="InvalidOperationException">The source document has no theme.</exception>
    public void Import(WordDocument source) {
        ArgumentNullException.ThrowIfNull(source);
        ThemePart sourcePart = source._wordprocessingDocument?.MainDocumentPart?.ThemePart
            ?? throw new InvalidOperationException("The source document has no theme.");
        ThemePart targetPart = EnsureThemePart();
        var copy = (A.Theme)(sourcePart.Theme ?? throw new InvalidOperationException("The source document has no theme.")).CloneNode(true);
        // Theme fills can embed pictures; bring each picture part along under a relationship of the target part.
        foreach (A.Blip blip in copy.Descendants<A.Blip>()) {
            if (blip.Embed?.Value is { } id && sourcePart.TryGetPartById(id, out OpenXmlPart? picture) && picture is ImagePart image) {
                ImagePart targetImage = targetPart.AddImagePart(image.ContentType);
                using (Stream stream = image.GetStream(FileMode.Open, FileAccess.Read)) targetImage.FeedData(stream);
                blip.Embed = targetPart.GetIdOfPart(targetImage);
            } else {
                blip.Embed = null;
            }
        }
        targetPart.Theme = copy;
        RefreshFallbacks();
    }

    private void EnsureEditable() {
        if (_document.FileOpenAccess == FileAccess.Read) throw new InvalidOperationException("A read-only document's theme cannot be changed.");
    }

    private ThemePart EnsureThemePart() {
        EnsureEditable();
        MainDocumentPart main = _document._wordprocessingDocument.MainDocumentPart
            ?? throw new InvalidOperationException("The document has no main part.");
        ThemePart part = main.ThemePart ?? main.AddNewPart<ThemePart>();
        if (part.Theme == null) WordDocument.GenerateThemePart1Content(part);
        return part;
    }

    private A.Theme EnsureTheme() {
        A.Theme theme = EnsureThemePart().Theme!;
        theme.ThemeElements ??= new A.ThemeElements();
        return theme;
    }

    private static string SlotToken(WordThemeColor color) => color switch {
        WordThemeColor.Dark1 or WordThemeColor.Text1 => "dk1",
        WordThemeColor.Light1 or WordThemeColor.Background1 => "lt1",
        WordThemeColor.Dark2 or WordThemeColor.Text2 => "dk2",
        WordThemeColor.Light2 or WordThemeColor.Background2 => "lt2",
        WordThemeColor.Accent1 => "accent1",
        WordThemeColor.Accent2 => "accent2",
        WordThemeColor.Accent3 => "accent3",
        WordThemeColor.Accent4 => "accent4",
        WordThemeColor.Accent5 => "accent5",
        WordThemeColor.Accent6 => "accent6",
        WordThemeColor.Hyperlink => "hlink",
        WordThemeColor.FollowedHyperlink => "folHlink",
        _ => throw new ArgumentOutOfRangeException(nameof(color), color, "Select a colour scheme slot.")
    };

    /// <summary>Rewrites the literal colours and fonts stored beside theme selections to the current theme values.</summary>
    private void RefreshFallbacks() {
        MainDocumentPart? main = _document._wordprocessingDocument?.MainDocumentPart;
        if (main == null) return;
        IEnumerable<OpenXmlElement> roots = _document.Styles.EnumerateContentRoots();
        if (main.StyleDefinitionsPart?.Styles is { } styles) roots = roots.Prepend(styles);
        foreach (OpenXmlElement root in roots) {
            foreach (OpenXmlElement element in root.Descendants()) {
                if (element is W.RunFonts fonts) {
                    RefreshFont(fonts);
                } else if (element.NamespaceUri == WordNamespace && element.HasAttributes) {
                    RefreshColor(element, "themeColor", "themeTint", "themeShade", element.LocalName == "color" ? "val" : "color");
                    RefreshColor(element, "themeFill", "themeFillTint", "themeFillShade", "fill");
                }
            }
        }
    }

    private void RefreshColor(OpenXmlElement element, string selector, string tint, string shade, string target) {
        string? theme = Attribute(element, selector);
        if (theme == null) return;
        string? resolved = WordRunFormatting.ResolveThemeColorHex(theme, Attribute(element, tint), Attribute(element, shade), _document);
        if (resolved != null) element.SetAttribute(new OpenXmlAttribute("w", target, WordNamespace, resolved));
    }

    private static string? Attribute(OpenXmlElement element, string localName) {
        foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
            if (attribute.LocalName == localName && attribute.NamespaceUri == WordNamespace) return attribute.Value;
        }
        return null;
    }

    private void RefreshFont(W.RunFonts fonts) {
        string? Typeface(EnumValue<W.ThemeFontValues>? selector) {
            string? token = selector?.InnerText;
            if (token == null) return null;
            // Only Latin typefaces are tracked; East Asian and complex-script slots keep their stored values.
            if (!token.EndsWith("HAnsi", StringComparison.Ordinal) && !token.EndsWith("Ascii", StringComparison.Ordinal)) return null;
            return GetFont(token.StartsWith("major", StringComparison.Ordinal) ? WordThemeFont.Major : WordThemeFont.Minor);
        }
        if (fonts.Ascii != null && Typeface(fonts.AsciiTheme) is { } ascii) fonts.Ascii = ascii;
        if (fonts.HighAnsi != null && Typeface(fonts.HighAnsiTheme) is { } highAnsi) fonts.HighAnsi = highAnsi;
    }
}
