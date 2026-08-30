using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static PdfCore.PdfTextRun CreatePdfTableCellTextRun(
        PptCore.PowerPointTableCell cell,
        A.Paragraph paragraph,
        A.TextCharacterPropertiesType? directProperties,
        string text,
        string? fallbackFontFamily,
        int slideNumber,
        PowerPointPdfSaveOptions options) {
        IReadOnlyList<A.TextCharacterPropertiesType> sources = ResolveTableTextPropertySources(cell, paragraph, directProperties);
        string? fontFamily = ReadRunFontName(cell, sources) ?? cell.FontName ?? fallbackFontFamily;
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

        return new PdfCore.PdfTextRun(
            text,
            bold: resolvedBold ?? cell.Bold,
            underline: underline.HasValue && underline.Value != TextUnderlineValues.None,
            color: ReadRunColor(cell, sources) ?? ParsePdfColor(cell.Color),
            italic: resolvedItalic ?? cell.Italic,
            strike: strike.HasValue && strike.Value != A.TextStrikeValues.NoStrike,
            fontSize: ReadRunFontSize(sources) ?? cell.FontSize,
            font: MapFont(fontFamily),
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
        A.TextCharacterPropertiesType? directProperties) {
        var sources = new List<A.TextCharacterPropertiesType>();
        if (directProperties != null) sources.Add(directProperties);

        A.DefaultRunProperties? paragraphDefaults = paragraph.ParagraphProperties?
            .GetFirstChild<A.DefaultRunProperties>();
        if (paragraphDefaults != null) sources.Add(paragraphDefaults);

        int level = paragraph.ParagraphProperties?.Level?.Value ?? 0;
        A.TextBody? textBody = paragraph.Ancestors<A.TextBody>().FirstOrDefault();
        sources.AddRange(FindDefaultRunProperties(textBody?.ListStyle, level));

        OpenXmlCompositeElement? otherStyle = cell.SlidePart?
            .SlideLayoutPart?
            .SlideMasterPart?
            .SlideMaster?
            .TextStyles?
            .OtherStyle;
        sources.AddRange(FindDefaultRunProperties(otherStyle, level));
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

    private static PdfCore.PdfColor? ReadRunColor(
        PptCore.PowerPointTableCell cell,
        IEnumerable<A.TextCharacterPropertiesType> sources) {
        A.ColorScheme? colorScheme = ResolveTableTextColorScheme(cell);
        foreach (A.TextCharacterPropertiesType source in sources) {
            A.SolidFill? fill = source.GetFirstChild<A.SolidFill>();
            if (fill == null) continue;
            A.SolidFill effectiveFill = ApplyTableTextColorMapping(cell, fill);
            OfficeIMO.Drawing.OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(effectiveFill, colorScheme);
            return color.HasValue ? PdfCore.PdfColor.FromOfficeColorOrNull(color.Value) : null;
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
        A.FontScheme? fontScheme = ResolveTableTextFontScheme(cell);
        foreach (A.TextCharacterPropertiesType source in sources) {
            string? font = source.GetFirstChild<A.LatinFont>()?.Typeface;
            if (string.IsNullOrWhiteSpace(font)) continue;
            if (font!.StartsWith("+mj-", StringComparison.OrdinalIgnoreCase)) {
                return fontScheme?.MajorFont?.LatinFont?.Typeface?.Value;
            }
            if (font.StartsWith("+mn-", StringComparison.OrdinalIgnoreCase)) {
                return fontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
            }
            return font;
        }
        return null;
    }

    private static A.FontScheme? ResolveTableTextFontScheme(PptCore.PowerPointTableCell cell) =>
        cell.SlidePart?.ThemeOverridePart?.ThemeOverride?.FontScheme
        ?? cell.SlidePart?.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.FontScheme
        ?? cell.SlidePart?.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;

    private static A.ColorScheme? ResolveTableTextColorScheme(PptCore.PowerPointTableCell cell) =>
        cell.SlidePart?.ThemeOverridePart?.ThemeOverride?.ColorScheme
        ?? cell.SlidePart?.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.ColorScheme
        ?? cell.SlidePart?.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;

    private static A.SolidFill ApplyTableTextColorMapping(
        PptCore.PowerPointTableCell cell,
        A.SolidFill fill) {
        A.SchemeColor? sourceScheme = fill.GetFirstChild<A.SchemeColor>();
        string sourceValue = sourceScheme?.GetAttribute("val", string.Empty).Value ?? string.Empty;
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
            : mapping?.GetAttribute(mappingSlot, string.Empty).Value ?? string.Empty;
        if (!IsMappedSchemeColor(target)) return fill;

        var clone = (A.SolidFill)fill.CloneNode(true);
        A.SchemeColor? clonedScheme = clone.GetFirstChild<A.SchemeColor>();
        if (clonedScheme != null) {
            clonedScheme.SetAttribute(new OpenXmlAttribute("val", string.Empty, target));
        }
        return clone;
    }

    private static OpenXmlElement? ResolveTableTextColorMapping(PptCore.PowerPointTableCell cell) {
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
