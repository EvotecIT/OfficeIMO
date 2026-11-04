using System.Collections.Generic;
using OfficeIMO.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeResolvedTextStyle(
            bool Bold,
            OfficeTextDecorationStyle UnderlineStyle,
            bool Italic,
            OfficeTextDecorationStyle StrikeStyle,
            bool AllCaps,
            PdfCore.PdfTextBaseline Baseline,
            double? FontSize,
            PdfCore.PdfStandardFont? Font,
            string? FontFamily,
            global::OfficeIMO.OfficeColor? Color,
            global::OfficeIMO.OfficeColor? BackgroundColor) {
            public bool SmallCaps { get; init; }
            public NativeTextSpacing TextSpacing { get; init; }
            public NativeTextSpacing ListMarkerTextSpacing { get; init; }
            public double EffectiveFontSize { get; init; }
            public PdfCore.PdfTextRun ApplyTo(PdfCore.PdfTextRun run) => TextSpacing.ApplyTo(run, EffectiveFontSize);
            internal bool Underline => UnderlineStyle != OfficeTextDecorationStyle.None;
            internal bool Strike => StrikeStyle != OfficeTextDecorationStyle.None;
        }

        private static void AddNativeText(
            PdfCore.PdfParagraph builder,
            string text,
            WordParagraph paragraph,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(paragraph, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            AddNativeCasedText(builder, text, style, tabStops, ref tabIndex);
            ResetNativeTextStyle(builder);
        }

        private static string ApplyNativeTextTransform(string text, WordParagraph paragraph, WordParagraph? fallback = null, NativeTableRunStyleDefaults tableRunStyleDefaults = default, NativeDocumentDefaults? nativeDefaults = null, NativeFontMap? nativeFontMap = null) {
            text = WordPaperText.NormalizeIdentifierBreaksForLayout(text);
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(paragraph, fallback, tableRunStyleDefaults, nativeDefaults, nativeFontMap);
            return style.AllCaps || style.SmallCaps ? text.ToUpperInvariant() : text;
        }

        private static void ApplyNativeTextStyle(PdfCore.PdfParagraph builder, WordParagraph paragraph, WordParagraph? fallback = null, NativeDocumentDefaults? nativeDefaults = null, NativeFontMap? nativeFontMap = null) =>
            ApplyNativeTextStyle(builder, ResolveNativeTextRunStyle(paragraph, fallback, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap));

        private static void ApplyNativeTextStyle(PdfCore.PdfParagraph builder, NativeResolvedTextStyle style) {
            builder.RunDefaults.HorizontalTextScaling = style.TextSpacing.WidthPercentage ?? 100D;
            builder.RunDefaults.CharacterSpacing = global::OfficeIMO.OfficeLength.Points(style.TextSpacing.CharacterSpacing ?? 0D);
            builder.RunDefaults.FeatureSettings = style.TextSpacing.ResolveFeatureSettings(style.EffectiveFontSize);
            builder.RunDefaults.Bold = style.Bold;
            builder.RunDefaults.Italic = style.Italic;
            builder.RunDefaults.UnderlineStyle = style.UnderlineStyle;
            builder.RunDefaults.StrikeStyle = style.StrikeStyle;
            builder.RunDefaults.Baseline = style.Baseline;
            if (style.FontSize.HasValue) {
                builder.RunDefaults.FontSize = global::OfficeIMO.OfficeLength.Points(style.FontSize.Value);
            }

            if (!string.IsNullOrWhiteSpace(style.FontFamily)) {
                builder.RunDefaults.FontFamily = style.FontFamily!; builder.RunDefaults.StandardFont = null;
            } else if (style.Font.HasValue) {
                builder.RunDefaults.StandardFont = style.Font.Value; builder.RunDefaults.FontFamily = null;
            }

            if (style.Color.HasValue) {
                builder.RunDefaults.Color = style.Color.Value;
            }

            if (style.BackgroundColor.HasValue) {
                builder.RunDefaults.BackgroundColor = style.BackgroundColor.Value;
            }
        }

        private static NativeResolvedTextStyle ResolveNativeTextRunStyle(WordParagraph paragraph, WordParagraph? fallback = null, NativeTableRunStyleDefaults tableRunStyleDefaults = default, NativeDocumentDefaults? nativeDefaults = null, NativeFontMap? nativeFontMap = null) {
            WordParagraph styleSource = fallback ?? paragraph;
            NativeDocumentDefaults resolvedNativeDefaults = nativeDefaults ?? GetNativeDocumentDefaults(styleSource._document);
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(styleSource);
            W.RunProperties? runProperties = GetNativeRunProperties(paragraph);
            NativeCharacterStyleDefaults characterStyleDefaults = GetNativeCharacterStyleDefaults(paragraph._document, runProperties);
            W.ParagraphMarkRunProperties? markerProperties = styleSource._paragraph?.ParagraphProperties?.ParagraphMarkRunProperties;
            NativeCharacterStyleDefaults markerCharacterStyle = GetNativeCharacterStyleDefaults(styleSource._document, markerProperties);

            bool bold = ReadNativeOnOff(runProperties?.GetFirstChild<W.Bold>()) ?? characterStyleDefaults.Bold ?? styleDefaults.Bold ?? tableRunStyleDefaults.Bold ?? resolvedNativeDefaults.RunStyle.Bold ?? false;
            bool italic = ReadNativeOnOff(runProperties?.GetFirstChild<W.Italic>()) ?? characterStyleDefaults.Italic ?? styleDefaults.Italic ?? tableRunStyleDefaults.Italic ?? resolvedNativeDefaults.RunStyle.Italic ?? false;
            OfficeTextDecorationStyle? directUnderlineStyle = MapNativeUnderlineStyle(runProperties?.GetFirstChild<W.Underline>());
            OfficeTextDecorationStyle underlineStyle = directUnderlineStyle ??
                characterStyleDefaults.UnderlineStyle ??
                styleDefaults.UnderlineStyle ??
                tableRunStyleDefaults.UnderlineStyle ??
                resolvedNativeDefaults.RunStyle.UnderlineStyle ??
                OfficeTextDecorationStyle.None;
            OfficeTextDecorationStyle strikeStyle = MapNativeStrikeStyle(runProperties) ??
                characterStyleDefaults.StrikeStyle ??
                styleDefaults.StrikeStyle ??
                tableRunStyleDefaults.StrikeStyle ??
                resolvedNativeDefaults.RunStyle.StrikeStyle ??
                OfficeTextDecorationStyle.None;
            bool allCaps =
                ReadNativeOnOff(runProperties?.GetFirstChild<W.Caps>()) ??
                characterStyleDefaults.AllCaps ??
                styleDefaults.AllCaps ??
                tableRunStyleDefaults.AllCaps ??
                resolvedNativeDefaults.RunStyle.AllCaps ??
                false;
            bool smallCaps =
                ReadNativeOnOff(runProperties?.GetFirstChild<W.SmallCaps>()) ??
                characterStyleDefaults.SmallCaps ??
                styleDefaults.SmallCaps ??
                tableRunStyleDefaults.SmallCaps ??
                resolvedNativeDefaults.RunStyle.SmallCaps ??
                false;
            PdfCore.PdfTextBaseline baseline = MapNativeTextBaseline(
                runProperties?.GetFirstChild<W.VerticalTextAlignment>()?.Val?.Value ??
                characterStyleDefaults.Baseline ??
                styleDefaults.Baseline ??
                tableRunStyleDefaults.Baseline ?? resolvedNativeDefaults.RunStyle.Baseline);
            double? fontSize = paragraph.NativeFontSizePoints.HasValue && paragraph.NativeFontSizePoints.Value > 0
                ? paragraph.NativeFontSizePoints.Value
                : characterStyleDefaults.FontSize ?? styleDefaults.FontSize ?? tableRunStyleDefaults.FontSize;
            fontSize = ResolveNativeComplexScriptFontSize(fontSize, runProperties, characterStyleDefaults, styleDefaults, tableRunStyleDefaults, resolvedNativeDefaults);
            // The paragraph wrapper's first run supplies no formatting inheritance to sibling runs.
            PdfCore.PdfStandardFont? font = ResolveNativeTextRunFont(paragraph, characterStyleDefaults, styleDefaults, tableRunStyleDefaults, resolvedNativeDefaults, nativeFontMap);
            string? fontFamily = ResolveNativeTextRunFontFamily(paragraph, characterStyleDefaults, styleDefaults, tableRunStyleDefaults, resolvedNativeDefaults, nativeFontMap);

            global::OfficeIMO.OfficeColor? color = TryGetNativeRunColor(runProperties, out global::OfficeIMO.OfficeColor? directColor)
                ? directColor
                : ResolveNativeInheritedRunColor(characterStyleDefaults, styleDefaults, tableRunStyleDefaults, resolvedNativeDefaults);
            color = ResolveNativeGroupedHeaderFooterTextColor(paragraph, color);
            global::OfficeIMO.OfficeColor? background = MapNativeHighlight(
                WordRunFormatting.ResolveRunHighlight(paragraph, tableRunStyleDefaults.Highlight));
            background ??= ParseNativeColor(WordRunFormatting.ResolveRunShadingFillColorHex(paragraph));

            return new NativeResolvedTextStyle(bold, underlineStyle, italic, strikeStyle, allCaps, baseline, fontSize, font, fontFamily, color, background) {
                SmallCaps = smallCaps,
                EffectiveFontSize = fontSize ?? nativeFontMap?.DefaultFontSize ?? resolvedNativeDefaults.FontSize,
                TextSpacing = resolvedNativeDefaults.TextSpacing.Merge(tableRunStyleDefaults.TextSpacing)
                    .Merge(styleDefaults.TextSpacing).Merge(characterStyleDefaults.TextSpacing).Merge(runProperties).Resolve(resolvedNativeDefaults.OpenTypeFeaturesEnabled),
                // Numbering uses the paragraph mark, independently of the first body run.
                ListMarkerTextSpacing = resolvedNativeDefaults.TextSpacing.Merge(tableRunStyleDefaults.TextSpacing)
                    .Merge(styleDefaults.TextSpacing).Merge(markerCharacterStyle.TextSpacing).Merge(markerProperties).Resolve(resolvedNativeDefaults.OpenTypeFeaturesEnabled)
            };
        }

        private static void ResetNativeTextStyle(PdfCore.PdfParagraph builder) {
            builder.RunDefaults.HorizontalTextScaling = 100D;
            builder.RunDefaults.CharacterSpacing = global::OfficeIMO.OfficeLength.Points(0D);
            builder.RunDefaults.FeatureSettings = global::OfficeIMO.Drawing.OfficeTextFeatureSettings.Default;
            builder.RunDefaults.Bold = false;
            builder.RunDefaults.Italic = false;
            builder.RunDefaults.UnderlineStyle = global::OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
            builder.RunDefaults.StrikeStyle = global::OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
            builder.RunDefaults.Baseline = PdfCore.PdfTextBaseline.Normal;
            builder.RunDefaults.Color = null;
            builder.RunDefaults.FontSize = null;
            builder.RunDefaults.StandardFont = null;
            builder.RunDefaults.FontFamily = null;
            builder.RunDefaults.BackgroundColor = null;
        }

    }
}
