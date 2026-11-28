using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordSettings {
    /// <summary>Gets or sets the document's default font size.</summary>
    public OfficeLength? FontSize {
        get => ReadFontSize("sz");
        set => SetFontSize("sz", value);
    }

    /// <summary>Gets or sets the default font size for complex-script text.</summary>
    public OfficeLength? FontSizeComplexScript {
        get => ReadFontSize("szCs");
        set => SetFontSize("szCs", value);
    }

    /// <summary>Gets the default ASCII font family or sets all four font families.</summary>
    public string? FontFamily {
        get => ReadFontFamily("ascii");
        set {
            OfficeXmlNode fonts = EnsureFonts();
            SetFontFamily(fonts, "ascii", "asciiTheme", value);
            SetFontFamily(fonts, "hAnsi", "hAnsiTheme", value);
            SetFontFamily(fonts, "eastAsia", "eastAsiaTheme", value);
            SetFontFamily(fonts, "cs", "cstheme", value);
            _document.MarkDocumentDefaultsEdited();
        }
    }

    /// <summary>Gets or sets the default High ANSI font family.</summary>
    public string? FontFamilyHighAnsi {
        get => ReadFontFamily("hAnsi");
        set => SetFontFamily("hAnsi", "hAnsiTheme", value);
    }

    /// <summary>Gets or sets the default East Asian font family.</summary>
    public string? FontFamilyEastAsia {
        get => ReadFontFamily("eastAsia");
        set => SetFontFamily("eastAsia", "eastAsiaTheme", value);
    }

    /// <summary>Gets or sets the default complex-script font family.</summary>
    public string? FontFamilyComplexScript {
        get => ReadFontFamily("cs");
        set => SetFontFamily("cs", "cstheme", value);
    }

    /// <summary>Gets or sets the document's default language.</summary>
    public string? Language {
        get => ReadDefaultRunProperties().ChildVal("lang");
        set {
            EnsureRunProperty("lang").SetAttribute(W, "val", value);
            _document.MarkDocumentDefaultsEdited();
        }
    }

    private OfficeXmlNode ReadDefaultRunProperties() => _document.GetDefaultRunProperties(edit: false);

    private OfficeLength? ReadFontSize(string name) => ReadDefaultRunProperties().ChildVal(name) is { } value
        ? OfficeLength.Emus((long)int.Parse(value, CultureInfo.InvariantCulture) * 6350)
        : null;

    private void SetFontSize(string name, OfficeLength? value) {
        string? size = value.HasValue ? checked((int)value.Value.ToHalfPoints()).ToString(CultureInfo.InvariantCulture) : null;
        EnsureRunProperty(name).SetAttribute(W, "val", size);
        _document.MarkDocumentDefaultsEdited();
    }

    private string? ReadFontFamily(string name) => ReadDefaultRunProperties().WChild("rFonts").WAttr(name);

    private void SetFontFamily(string name, string themeName, string? value) {
        SetFontFamily(EnsureFonts(), name, themeName, string.IsNullOrEmpty(value) ? null : value);
        _document.MarkDocumentDefaultsEdited();
    }

    private static void SetFontFamily(OfficeXmlNode fonts, string name, string themeName, string? value) {
        fonts.SetAttribute(W, name, value);
        fonts.RemoveAttribute(W, themeName);
    }

    private OfficeXmlNode EnsureFonts() {
        OfficeXmlNode properties = RequireDefaultRunProperties();
        OfficeXmlNode fonts = properties.WChild("rFonts");
        if (fonts.Exists) return fonts;
        fonts = WordSettingsXml.AddChild(properties, "rFonts", WRPrView.Order);
        fonts.SetAttribute(W, "asciiTheme", "minorHAnsi");
        fonts.SetAttribute(W, "hAnsiTheme", "minorHAnsi");
        fonts.SetAttribute(W, "eastAsiaTheme", "minorHAnsi");
        fonts.SetAttribute(W, "cstheme", "minorBidi");
        return fonts;
    }

    private OfficeXmlNode EnsureRunProperty(string name) => WordSettingsXml.EnsureChild(RequireDefaultRunProperties(), name, WRPrView.Order);

    private OfficeXmlNode RequireDefaultRunProperties() {
        OfficeXmlNode properties = _document.GetDefaultRunProperties(edit: true);
        return properties.Exists ? properties : throw new InvalidOperationException("Document default styles are unavailable.");
    }
}
