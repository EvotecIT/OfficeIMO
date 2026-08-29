using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal readonly struct PowerPointEffectiveRunStyle {
        internal PowerPointEffectiveRunStyle(
            bool? bold,
            bool? italic,
            PowerPointUnderlineStyle? underlineStyle,
            PowerPointStrikeStyle? strikeStyle,
            PowerPointCapitalization? capitalization,
            double? baselinePercent,
            double? fontSizePoints,
            string? fontName,
            string? color,
            string? language) {
            Bold = bold;
            Italic = italic;
            UnderlineStyle = underlineStyle;
            StrikeStyle = strikeStyle;
            Capitalization = capitalization;
            BaselinePercent = baselinePercent;
            FontSizePoints = fontSizePoints;
            FontName = fontName;
            Color = color;
            Language = language;
        }

        internal bool? Bold { get; }
        internal bool? Italic { get; }
        internal PowerPointUnderlineStyle? UnderlineStyle { get; }
        internal PowerPointStrikeStyle? StrikeStyle { get; }
        internal PowerPointCapitalization? Capitalization { get; }
        internal double? BaselinePercent { get; }
        internal double? FontSizePoints { get; }
        internal string? FontName { get; }
        internal string? Color { get; }
        internal string? Language { get; }
    }

    internal static class PowerPointEffectiveRunStyleResolver {
        internal static PowerPointEffectiveRunStyle Resolve(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) {
            IReadOnlyList<A.TextCharacterPropertiesType> sources = ResolveSources(run, paragraph, listStyle, masterTextStyle);
            bool? bold = sources.Select(source => source.Bold == null ? (bool?)null : source.Bold.Value)
                .FirstOrDefault(value => value.HasValue);
            bool? italic = sources.Select(source => source.Italic == null ? (bool?)null : source.Italic.Value)
                .FirstOrDefault(value => value.HasValue);
            A.TextUnderlineValues? underline = sources.Select(source => source.Underline?.Value)
                .FirstOrDefault(value => value.HasValue);
            A.TextStrikeValues? strike = sources.Select(source => source.Strike?.Value)
                .FirstOrDefault(value => value.HasValue);
            A.TextCapsValues? capitalization = sources.Select(source => source.Capital?.Value)
                .FirstOrDefault(value => value.HasValue);
            int? baseline = sources.Select(source => source.Baseline?.Value)
                .FirstOrDefault(value => value.HasValue);
            int? fontSize = sources.Select(source => source.FontSize?.Value)
                .FirstOrDefault(value => value.HasValue);
            string? fontName = sources.Select(source => source.GetFirstChild<A.LatinFont>()?.Typeface?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            string? color = sources.Select(source => source.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            string? language = sources.Select(source => source.Language?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

            return new PowerPointEffectiveRunStyle(
                bold,
                italic,
                underline?.ToOfficeEnum(),
                ToStrikeStyle(strike),
                ToCapitalization(capitalization),
                baseline.HasValue ? baseline.Value / 1000D : (double?)null,
                fontSize.HasValue ? fontSize.Value / 100D : (double?)null,
                fontName,
                color,
                language);
        }

        private static IReadOnlyList<A.TextCharacterPropertiesType> ResolveSources(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) {
            var sources = new List<A.TextCharacterPropertiesType>();
            if (run.RunProperties != null) sources.Add(run.RunProperties);
            A.DefaultRunProperties? paragraphDefaults = paragraph.Paragraph.ParagraphProperties?
                .GetFirstChild<A.DefaultRunProperties>();
            if (paragraphDefaults != null) sources.Add(paragraphDefaults);
            int level = paragraph.Paragraph.ParagraphProperties?.Level?.Value ?? 0;
            sources.AddRange(FindDefaultRunProperties(listStyle, level));
            sources.AddRange(FindDefaultRunProperties(masterTextStyle, level));
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

        private static PowerPointStrikeStyle? ToStrikeStyle(A.TextStrikeValues? value) {
            if (!value.HasValue) return null;
            if (value.Value == A.TextStrikeValues.NoStrike) return PowerPointStrikeStyle.None;
            if (value.Value == A.TextStrikeValues.SingleStrike) return PowerPointStrikeStyle.Single;
            if (value.Value == A.TextStrikeValues.DoubleStrike) return PowerPointStrikeStyle.Double;
            return null;
        }

        private static PowerPointCapitalization? ToCapitalization(A.TextCapsValues? value) {
            if (!value.HasValue) return null;
            if (value.Value == A.TextCapsValues.None) return PowerPointCapitalization.None;
            if (value.Value == A.TextCapsValues.Small) return PowerPointCapitalization.SmallCaps;
            if (value.Value == A.TextCapsValues.All) return PowerPointCapitalization.AllCaps;
            return null;
        }
    }
}
