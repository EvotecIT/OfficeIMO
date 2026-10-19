using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Reusable portable browser-font profile for OfficeIMO HTML, image, and PDF rendering.
/// Hosts supply the pinned font bytes; this type owns the family aliases, fallback order,
/// PDF embedding, and render-option activation.
/// </summary>
public sealed class HtmlPortableBrowserFontProfile {
    private static readonly string[] PortableDefaultAliases = [
        "Arial", "Helvetica", "Calibri", "Aptos", "Segoe UI", "Tahoma", "Verdana",
        "sans", "sans-serif", "ui-sans-serif", "system-ui", "-apple-system", "BlinkMacSystemFont",
        "Times", "Times Roman", "Times-Roman", "Times New Roman", "serif",
        "Courier", "Courier New", "monospace", "ui-monospace"
    ];
    private static readonly string[] PortableSymbolAliases = ["Symbol", "ZapfDingbats", "Zapf Dingbats"];
    private readonly PdfEmbeddedFontFamily _defaultPdfFontFamily;
    private readonly PdfEmbeddedFontFallbackSet _pdfFontFallbacks;

    private HtmlPortableBrowserFontProfile(string id, OfficeFontFallbackPack fontPack,
        PdfEmbeddedFontFamily defaultPdfFontFamily, PdfEmbeddedFontFallbackSet pdfFontFallbacks,
        IReadOnlyList<PdfFontFamilySubstitution> substitutions) {
        Id = id;
        FontPack = fontPack;
        _defaultPdfFontFamily = defaultPdfFontFamily;
        _pdfFontFallbacks = pdfFontFallbacks;
        FontFamilySubstitutions = substitutions;
    }

    /// <summary>Default Latin Office document family.</summary>
    public const string DefaultFontFamily = "Carlito";
    /// <summary>Bounded Japanese fallback family.</summary>
    public const string JapaneseFallbackFontFamily = "OfficeIMO Japanese Common";
    /// <summary>Arabic fallback family.</summary>
    public const string ArabicFallbackFontFamily = "Noto Sans Arabic";
    /// <summary>Symbol fallback family.</summary>
    public const string SymbolFallbackFontFamily = "Noto Sans Symbols 2";
    /// <summary>Default CSS family sequence used by the portable profile.</summary>
    public const string DefaultLayoutFontFamilies =
        "Carlito, 'OfficeIMO Japanese Common', 'Noto Sans Arabic', 'Noto Sans Symbols 2'";

    /// <summary>Stable source package identifier.</summary>
    public string Id { get; }
    /// <summary>Immutable drawing fallback pack shared by HTML, image, and PDF rendering.</summary>
    public OfficeFontFallbackPack FontPack { get; }
    /// <summary>Lowercase SHA-256 identity of the activated faces, aliases, and fallback order.</summary>
    public string Fingerprint => FontPack.Fingerprint;
    /// <summary>Explicit PDF family substitutions carried by the portable profile.</summary>
    public IReadOnlyList<PdfFontFamilySubstitution> FontFamilySubstitutions { get; }

    /// <summary>
    /// Creates the standard seven-face portable profile from a caller-owned asset reader.
    /// The reader must return the exact named font file bytes or throw.
    /// </summary>
    public static HtmlPortableBrowserFontProfile Create(string id, Func<string, byte[]> readAsset) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("A portable font-package id is required.", nameof(id));
        if (readAsset == null) throw new ArgumentNullException(nameof(readAsset));
        byte[] regular = Read(readAsset, "Carlito-Regular.ttf");
        byte[] bold = Read(readAsset, "Carlito-Bold.ttf");
        byte[] italic = Read(readAsset, "Carlito-Italic.ttf");
        byte[] boldItalic = Read(readAsset, "Carlito-BoldItalic.ttf");
        byte[] japanese = Read(readAsset, "NotoSansJP-OfficeIMO-Common.ttf");
        byte[] arabic = Read(readAsset, "NotoSansArabic-Regular.ttf");
        byte[] symbols = Read(readAsset, "NotoSansSymbols2-Regular.ttf");

        var fonts = new OfficeFontFaceCollection()
            .Add(DefaultFontFamily, regular, OfficeFontStyle.Regular)
            .Add(DefaultFontFamily, bold, OfficeFontStyle.Bold)
            .Add(DefaultFontFamily, italic, OfficeFontStyle.Italic)
            .Add(DefaultFontFamily, boldItalic, OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        foreach (string alias in PortableDefaultAliases) fonts.AddAlias(alias, DefaultFontFamily);
        fonts.Add(SymbolFallbackFontFamily, symbols);
        foreach (string alias in PortableSymbolAliases) fonts.AddAlias(alias, SymbolFallbackFontFamily);
        fonts.Add(JapaneseFallbackFontFamily, japanese)
            .Add(ArabicFallbackFontFamily, arabic)
            .AddFallbackFamily(JapaneseFallbackFontFamily)
            .AddFallbackFamily(ArabicFallbackFontFamily)
            .AddFallbackFamily(SymbolFallbackFontFamily);

        var substitutions = Array.AsReadOnly(new[] {
            new PdfFontFamilySubstitution("Calibri", DefaultFontFamily, PdfFontFamilySubstitutionImpact.Compatible),
            new PdfFontFamilySubstitution("Calibri Light", DefaultFontFamily, PdfFontFamilySubstitutionImpact.Compatible),
            new PdfFontFamilySubstitution("Aptos", DefaultFontFamily, PdfFontFamilySubstitutionImpact.LayoutSensitive),
            new PdfFontFamilySubstitution("Aptos Display", DefaultFontFamily, PdfFontFamilySubstitutionImpact.LayoutSensitive),
            new PdfFontFamilySubstitution("Symbol", SymbolFallbackFontFamily, PdfFontFamilySubstitutionImpact.LayoutSensitive)
        });
        return new HtmlPortableBrowserFontProfile(id.Trim(),
            new OfficeFontFallbackPack(id.Trim(), DefaultLayoutFontFamilies, fonts),
            new PdfEmbeddedFontFamily(DefaultFontFamily, regular, bold, italic, boldItalic),
            new PdfEmbeddedFontFallbackSet([
                new PdfEmbeddedFontFallbackCandidate(JapaneseFallbackFontFamily, japanese),
                new PdfEmbeddedFontFallbackCandidate(ArabicFallbackFontFamily, arabic),
                new PdfEmbeddedFontFallbackCandidate(SymbolFallbackFontFamily, symbols)
            ]), substitutions);
    }

    /// <summary>Creates an independent drawing-font collection with aliases and ordered fallbacks.</summary>
    public OfficeFontFaceCollection CreateDrawingFonts() => FontPack.Fonts;

    /// <summary>Creates deterministic PDF writer options with embedded fonts and explicit substitutions.</summary>
    public PdfOptions CreatePdfOptions(IOfficeTextShapingProvider? textShapingProvider = null) {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            HeaderFont = PdfStandardFont.Helvetica,
            FooterFont = PdfStandardFont.Helvetica,
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly,
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            TextShapingMode = PdfTextShapingMode.LatinLigatures
        }.SetTextShapingProvider(textShapingProvider);
        options.RegisterFontFamily(PdfStandardFont.Helvetica, _defaultPdfFontFamily);
        options.RegisterNamedFontFamily(_defaultPdfFontFamily);
        options.RegisterEmbeddedFontFallbacks(_pdfFontFallbacks);
        foreach (PdfFontFamilySubstitution substitution in FontFamilySubstitutions) {
            options.RegisterFontFamilySubstitution(
                substitution.SourceFontFamily, substitution.TargetFontFamily, substitution.Impact);
        }
        return options;
    }

    /// <summary>Creates HTML rendering options with the same fonts for layout, images, and PDF output.</summary>
    public HtmlToPdfOptions CreateHtmlOptions(IOfficeTextShapingProvider? textShapingProvider = null) {
        var options = new HtmlToPdfOptions {
            DefaultFontFamily = DefaultLayoutFontFamilies,
            PdfOptions = CreatePdfOptions(textShapingProvider),
            FontFamily = _defaultPdfFontFamily,
            TextShapingMode = PdfTextShapingMode.LatinLigatures,
            TextShapingProvider = textShapingProvider,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.UseRenderingProfile(FontPack.CreateRenderingProfile(textShapingProvider));
        options.DefaultFontFamily = DefaultLayoutFontFamilies;
        return options;
    }

    private static byte[] Read(Func<string, byte[]> readAsset, string name) {
        byte[] bytes = readAsset(name) ?? throw new InvalidOperationException("The portable font asset '" + name + "' is missing.");
        if (bytes.Length == 0) throw new InvalidOperationException("The portable font asset '" + name + "' is empty.");
        return bytes;
    }
}
