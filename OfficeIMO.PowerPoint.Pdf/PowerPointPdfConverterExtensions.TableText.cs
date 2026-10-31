using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    // Adapt the canonical table-region cascade to the same property/color/font
    // resolver used for direct runs. Local paragraph formatting wins over styles.
    private static IReadOnlyList<A.TextCharacterPropertiesType> CreateTableTextDefaults(
        PptCore.PowerPointTable table, int row, int column) {
        var defaults = new List<A.TextCharacterPropertiesType>();
        foreach (A.TableCellTextStyle style in PptCore.PowerPointSlideImageRenderer.ResolveTableCellTextStylesForExport(table, row, column)) {
            var properties = new A.DefaultRunProperties();
            string? bold = ReadOptionalTableTextAttribute(style, "b");
            string? italic = ReadOptionalTableTextAttribute(style, "i");
            if (!string.IsNullOrEmpty(bold) && bold != "def") properties.Bold = bold is "on" or "true" or "1";
            if (!string.IsNullOrEmpty(italic) && italic != "def") properties.Italic = italic is "on" or "true" or "1";

            OpenXmlElement? color = style.ChildElements.FirstOrDefault(child => child.LocalName is
                "srgbClr" or "scrgbClr" or "schemeClr" or "sysClr" or "prstClr" or "hslClr");
            if (color != null) properties.Append(new A.SolidFill(color.CloneNode(true)));
            string? index = ReadOptionalTableTextAttribute(style.GetFirstChild<A.FontReference>(), "idx");
            if (index is "major" or "minor") {
                properties.Append(new A.LatinFont { Typeface = index == "major" ? "+mj-lt" : "+mn-lt" });
            } else if (style.ChildElements.FirstOrDefault(child => child.LocalName == "font")?.GetFirstChild<A.LatinFont>() is A.LatinFont latin) {
                properties.Append(latin.CloneNode(true));
            }
            defaults.Add(properties);
        }
        return defaults;
    }

    // Open XML GetAttribute throws for absent attributes; style-region defaults
    // routinely omit b/i and optional color-map slots.
    private static string? ReadOptionalTableTextAttribute(OpenXmlElement? element, string name) =>
        element?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == name
            && string.IsNullOrEmpty(attribute.NamespaceUri)).Value;

    private static PdfCore.PdfTextRun CreatePdfTableCellTextRun(
        PptCore.PowerPointTableCell cell,
        A.Paragraph paragraph,
        A.TextCharacterPropertiesType? directProperties,
        string text,
        string? fallbackFontFamily,
        int slideNumber,
        PowerPointPdfExportOptions options,
        IReadOnlyList<A.TextCharacterPropertiesType>? tableDefaults = null,
        double? fallbackFontSize = null) {
        IReadOnlyList<A.TextCharacterPropertiesType> sources = ResolveTableTextPropertySources(cell, paragraph, directProperties, tableDefaults);
        int? characterSpacing = sources.Select(source => source.Spacing?.Value).FirstOrDefault(value => value.HasValue);
        if (!string.IsNullOrEmpty(text) && characterSpacing.GetValueOrDefault() != 0)
            AddWarning(options, slideNumber, "character-spacing-approximation",
                "PDF table text uses default character spacing; explicit spacing is retained in the presentation.");
        string? fontFamily = ReadRunFontName(cell, sources) ?? fallbackFontFamily;
        A.TextUnderlineValues? underline = sources
            .Select(source => source.Underline?.Value)
            .FirstOrDefault(value => value.HasValue);
        A.TextStrikeValues? strike = sources
            .Select(source => source.Strike?.Value)
            .FirstOrDefault(value => value.HasValue);
        int? baseline = sources
            .Select(source => source.Baseline?.Value)
            .FirstOrDefault(value => value.HasValue);
        A.TextCapsValues? capitalization = sources
            .Select(source => source.Capital?.Value)
            .FirstOrDefault(value => value.HasValue);
        string? language = sources
            .Select(source => source.Language?.Value)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (capitalization == A.TextCapsValues.All || capitalization == A.TextCapsValues.Small) {
            text = ApplyPowerPointDisplayCase(
                text,
                capitalization == A.TextCapsValues.Small
                    ? PptCore.PowerPointCapitalization.SmallCaps
                    : PptCore.PowerPointCapitalization.AllCaps,
                language,
                options,
                slideNumber);
        }

        bool? resolvedBold = sources
            .Select(source => source.Bold == null ? (bool?)null : source.Bold.Value)
            .FirstOrDefault(value => value.HasValue);
        bool? resolvedItalic = sources
            .Select(source => source.Italic == null ? (bool?)null : source.Italic.Value)
            .FirstOrDefault(value => value.HasValue);

        Uri? hyperlink = cell.OwnerPart == null ? null : PptCore.PowerPointHyperlinkResolver.Resolve(
            cell.OwnerPart, cell.SlidePart, directProperties?.GetFirstChild<A.HyperlinkOnClick>());
        string? destination = ResolveSlideLinkDestination(hyperlink, slideNumber, options, "table text run");

        return new PdfCore.PdfTextRun(
            text,
            bold: resolvedBold ?? false,
            underline: underline.HasValue && underline.Value != TextUnderlineValues.None,
            color: ReadRunColor(cell, sources),
            italic: resolvedItalic ?? false,
            strike: strike.HasValue && strike.Value != A.TextStrikeValues.NoStrike,
            fontSize: ((double?)(ReadRunFontSize(sources) ?? fallbackFontSize) is double physicalLength1 ? global::OfficeIMO.OfficeLength.Points(physicalLength1) : (global::OfficeIMO.OfficeLength?)null),
            font: MapFont(fontFamily),
            linkUri: hyperlink is { IsAbsoluteUri: true } ? hyperlink.AbsoluteUri : null,
            linkDestinationName: destination,
            baseline: MapPowerPointBaseline(baseline.HasValue ? baseline.Value / 1000D : (double?)null, options, slideNumber),
            fontFamily: fontFamily,
            underlineStyle: MapPowerPointUnderline(underline),
            strikeStyle: strike == A.TextStrikeValues.DoubleStrike
                ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Double
                : strike == A.TextStrikeValues.SingleStrike
                    ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single
                    : OfficeIMO.Drawing.OfficeTextDecorationStyle.None);
    }

    private static IReadOnlyList<A.TextCharacterPropertiesType> ResolveTableTextPropertySources(
        PptCore.PowerPointTableCell cell,
        A.Paragraph paragraph,
        A.TextCharacterPropertiesType? directProperties,
        IReadOnlyList<A.TextCharacterPropertiesType>? tableDefaults = null) {
        var sources = new List<A.TextCharacterPropertiesType>();
        if (directProperties != null) sources.Add(directProperties);

        A.DefaultRunProperties? paragraphDefaults = paragraph.ParagraphProperties?
            .GetFirstChild<A.DefaultRunProperties>();
        if (paragraphDefaults != null) sources.Add(paragraphDefaults);

        int level = paragraph.ParagraphProperties?.Level?.Value ?? 0;
        A.TextBody? textBody = paragraph.Ancestors<A.TextBody>().FirstOrDefault();
        sources.AddRange(FindDefaultRunProperties(textBody?.ListStyle, level));
        if (tableDefaults != null) sources.AddRange(tableDefaults);

        SlideMasterPart? master = ResolveTableMasterPart(cell);
        OpenXmlCompositeElement? otherStyle = master?.SlideMaster?.TextStyles?.OtherStyle;
        sources.AddRange(FindDefaultRunProperties(otherStyle, level));
        sources.AddRange(FindDefaultRunProperties((cell.SlidePart ?? (OpenXmlPart?)master)?.GetParentParts()
            .OfType<DocumentFormat.OpenXml.Packaging.PresentationPart>().FirstOrDefault()?
            .Presentation?.DefaultTextStyle, level));
        return sources;
    }

    private static IEnumerable<A.DefaultRunProperties> FindDefaultRunProperties(
        OpenXmlCompositeElement? container,
        int level) {
        A.DefaultRunProperties? levelDefaults = container?
            .ChildElements
            .OfType<A.TextParagraphPropertiesType>()
            .FirstOrDefault(properties => GetTextLevel(properties) == level)?
            .GetFirstChild<A.DefaultRunProperties>();
        if (levelDefaults != null) yield return levelDefaults;
        A.DefaultRunProperties? fallbackDefaults = container?
            .GetFirstChild<A.DefaultParagraphProperties>()?
            .GetFirstChild<A.DefaultRunProperties>();
        if (fallbackDefaults != null) yield return fallbackDefaults;
    }

    private static int GetTextLevel(A.TextParagraphPropertiesType properties) => properties switch {
        A.Level1ParagraphProperties => 0,
        A.Level2ParagraphProperties => 1,
        A.Level3ParagraphProperties => 2,
        A.Level4ParagraphProperties => 3,
        A.Level5ParagraphProperties => 4,
        A.Level6ParagraphProperties => 5,
        A.Level7ParagraphProperties => 6,
        A.Level8ParagraphProperties => 7,
        A.Level9ParagraphProperties => 8,
        _ => -1
    };

    private static global::OfficeIMO.OfficeColor? ReadRunColor(
        PptCore.PowerPointTableCell cell,
        IEnumerable<A.TextCharacterPropertiesType> sources) {
        A.ColorScheme? colorScheme = ResolveTableTextColorScheme(cell);
        foreach (A.TextCharacterPropertiesType source in sources) {
            A.SolidFill? fill = source.GetFirstChild<A.SolidFill>();
            if (fill == null) continue;
            A.SolidFill effectiveFill = ApplyTableTextColorMapping(cell, fill);
            OfficeIMO.OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(effectiveFill, colorScheme);
            return color.HasValue ? (color.Value.A == 0 ? (global::OfficeIMO.OfficeColor?)null : color.Value) : null;
        }
        return null;
    }

    private static double? ReadRunFontSize(IEnumerable<A.TextCharacterPropertiesType> sources) {
        foreach (A.TextCharacterPropertiesType source in sources) {
            int? size = source.FontSize?.Value;
            if (size.HasValue) return size.Value / 100D;
        }
        return null;
    }

    private static string? ReadRunFontName(
        PptCore.PowerPointTableCell cell,
        IEnumerable<A.TextCharacterPropertiesType> sources) {
        foreach (A.TextCharacterPropertiesType source in sources) {
            string? font = source.GetFirstChild<A.LatinFont>()?.Typeface;
            if (string.IsNullOrWhiteSpace(font)) continue;
            if (font!.StartsWith("+mj-", StringComparison.OrdinalIgnoreCase)) {
                return PptCore.PowerPointEffectiveRunStyleResolver.ResolveLatinThemeFontName(cell.SlidePart ?? cell.OwnerPart, major: true);
            }
            if (font.StartsWith("+mn-", StringComparison.OrdinalIgnoreCase)) {
                return PptCore.PowerPointEffectiveRunStyleResolver.ResolveLatinThemeFontName(cell.SlidePart ?? cell.OwnerPart);
            }
            return font;
        }
        return null;
    }

    private static string? ResolveTableDefaultFontFamily(PptCore.PowerPointTable table) =>
        PptCore.PowerPointEffectiveRunStyleResolver.ResolveLatinThemeFontName(table.OwnerSlide?.SlidePart ?? table.RelationshipOwnerPart);

    private static SlideMasterPart? ResolveTableMasterPart(PptCore.PowerPointTableCell cell) =>
        (cell.SlidePart ?? cell.OwnerPart) switch {
            SlidePart slide => slide.SlideLayoutPart?.SlideMasterPart,
            SlideLayoutPart layout => layout.SlideMasterPart,
            SlideMasterPart master => master,
            _ => null
        };

    private static A.ColorScheme? ResolveTableTextColorScheme(PptCore.PowerPointTableCell cell) =>
        PptCore.PowerPointEffectiveRunStyleResolver.ResolveColorScheme(cell.SlidePart ?? cell.OwnerPart);

    private static A.SolidFill ApplyTableTextColorMapping(
        PptCore.PowerPointTableCell cell,
        A.SolidFill fill) {
        A.SchemeColor? sourceScheme = fill.GetFirstChild<A.SchemeColor>();
        string sourceValue = ReadOptionalTableTextAttribute(sourceScheme, "val") ?? string.Empty;
        string? mappingSlot = sourceValue switch {
            "bg1" => "bg1",
            "tx1" => "tx1",
            "bg2" => "bg2",
            "tx2" => "tx2",
            "accent1" => "accent1",
            "accent2" => "accent2",
            "accent3" => "accent3",
            "accent4" => "accent4",
            "accent5" => "accent5",
            "accent6" => "accent6",
            "hlink" => "hlink",
            "folHlink" => "folHlink",
            _ => null
        };
        OpenXmlElement? mapping = ResolveTableTextColorMapping(cell);
        string target = mappingSlot == null
            ? string.Empty
            : ReadOptionalTableTextAttribute(mapping, mappingSlot) ?? string.Empty;
        if (!IsMappedSchemeColor(target)) return fill;

        var clone = (A.SolidFill)fill.CloneNode(true);
        A.SchemeColor? clonedScheme = clone.GetFirstChild<A.SchemeColor>();
        if (clonedScheme != null) {
            clonedScheme.SetAttribute(new OpenXmlAttribute("val", string.Empty, target));
        }
        return clone;
    }

    private static OpenXmlElement? ResolveTableTextColorMapping(PptCore.PowerPointTableCell cell) {
        if (cell.SlidePart == null) return PptCore.PowerPointEffectiveRunStyleResolver.ResolveColorMap(cell.OwnerPart);
        P.ColorMap? master = cell.SlidePart?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.ColorMap;
        P.ColorMapOverride? slideOverride = cell.SlidePart?.Slide?.ColorMapOverride;
        if (slideOverride?.GetFirstChild<A.OverrideColorMapping>() is A.OverrideColorMapping slideMapping) {
            return slideMapping;
        }
        if (slideOverride?.GetFirstChild<A.MasterColorMapping>() != null) return master;

        P.ColorMapOverride? layoutOverride = cell.SlidePart?.SlideLayoutPart?.SlideLayout?.ColorMapOverride;
        if (layoutOverride?.GetFirstChild<A.OverrideColorMapping>() is A.OverrideColorMapping layoutMapping) {
            return layoutMapping;
        }
        return master;
    }

    private static bool IsMappedSchemeColor(string value) => value is
        "dk1" or "lt1" or "dk2" or "lt2"
        or "accent1" or "accent2" or "accent3" or "accent4" or "accent5" or "accent6"
        or "hlink" or "folHlink";

    private static string? ReadRunFontName(A.RunProperties? properties) =>
        properties?.GetFirstChild<A.LatinFont>()?.Typeface;
}
