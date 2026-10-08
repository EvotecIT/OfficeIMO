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
            PdfCore.PdfColor? Color,
            PdfCore.PdfColor? BackgroundColor) {
            public NativeTextSpacing TextSpacing { get; init; }
            public NativeTextSpacing ListMarkerTextSpacing { get; init; }
            internal bool Underline => UnderlineStyle != OfficeTextDecorationStyle.None;
            internal bool Strike => StrikeStyle != OfficeTextDecorationStyle.None;
        }

        private static void AddNativeText(
            PdfCore.PdfParagraphBuilder builder,
            string text,
            WordParagraph paragraph,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            ApplyNativeTextStyle(builder, paragraph, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            AddNativeRunText(builder, ApplyNativeTextTransform(text, paragraph, nativeFontMap: nativeFontMap), tabStops, ref tabIndex);
            ResetNativeTextStyle(builder);
        }

        private static string ApplyNativeTextTransform(string text, WordParagraph paragraph, WordParagraph? fallback = null, NativeTableRunStyleDefaults tableRunStyleDefaults = default, NativeDocumentDefaults? nativeDefaults = null, NativeFontMap? nativeFontMap = null) =>
            ResolveNativeTextRunStyle(paragraph, fallback, tableRunStyleDefaults, nativeDefaults, nativeFontMap).AllCaps
                ? text.ToUpperInvariant()
                : text;

        private static void ApplyNativeTextStyle(PdfCore.PdfParagraphBuilder builder, WordParagraph paragraph, WordParagraph? fallback = null, NativeDocumentDefaults? nativeDefaults = null, NativeFontMap? nativeFontMap = null) =>
            ApplyNativeTextStyle(builder, ResolveNativeTextRunStyle(paragraph, fallback, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap));

        private static void ApplyNativeTextStyle(PdfCore.PdfParagraphBuilder builder, NativeResolvedTextStyle style) {
            builder.HorizontalTextScaling(style.TextSpacing.WidthPercentage ?? 100D);
            builder.CharacterSpacing(style.TextSpacing.CharacterSpacing ?? 0D);
            builder.Bold(style.Bold);
            builder.Italic(style.Italic);
            builder.Underline(style.UnderlineStyle);
            builder.Strike(style.StrikeStyle);
            builder.Baseline(style.Baseline);
            if (style.FontSize.HasValue) {
                builder.FontSize(style.FontSize.Value);
            }

            if (!string.IsNullOrWhiteSpace(style.FontFamily)) {
                builder.FontFamily(style.FontFamily!);
            } else if (style.Font.HasValue) {
                builder.Font(style.Font.Value);
            }

            if (style.Color.HasValue) {
                builder.Color(style.Color.Value);
            }

            if (style.BackgroundColor.HasValue) {
                builder.BackgroundColor(style.BackgroundColor.Value);
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

            bool bold = ReadNativeOnOff(runProperties?.GetFirstChild<W.Bold>()) ?? characterStyleDefaults.Bold ?? styleDefaults.Bold ?? tableRunStyleDefaults.Bold ?? false;
            bool italic = ReadNativeOnOff(runProperties?.GetFirstChild<W.Italic>()) ?? characterStyleDefaults.Italic ?? styleDefaults.Italic ?? tableRunStyleDefaults.Italic ?? false;
            OfficeTextDecorationStyle? directUnderlineStyle = MapNativeUnderlineStyle(runProperties?.GetFirstChild<W.Underline>());
            OfficeTextDecorationStyle underlineStyle = directUnderlineStyle ??
                characterStyleDefaults.UnderlineStyle ??
                styleDefaults.UnderlineStyle ??
                tableRunStyleDefaults.UnderlineStyle ??
                OfficeTextDecorationStyle.None;
            OfficeTextDecorationStyle strikeStyle = MapNativeStrikeStyle(runProperties) ??
                characterStyleDefaults.StrikeStyle ??
                styleDefaults.StrikeStyle ??
                tableRunStyleDefaults.StrikeStyle ??
                OfficeTextDecorationStyle.None;
            bool allCaps =
                ReadNativeOnOff(runProperties?.GetFirstChild<W.Caps>()) ??
                ReadNativeOnOff(runProperties?.GetFirstChild<W.SmallCaps>()) ??
                characterStyleDefaults.AllCaps ??
                styleDefaults.AllCaps ??
                tableRunStyleDefaults.AllCaps ??
                false;
            PdfCore.PdfTextBaseline baseline = MapNativeTextBaseline(
                runProperties?.GetFirstChild<W.VerticalTextAlignment>()?.Val?.Value ??
                characterStyleDefaults.Baseline ??
                styleDefaults.Baseline ??
                tableRunStyleDefaults.Baseline);
            double? fontSize = paragraph.FontSizePoints.HasValue && paragraph.FontSizePoints.Value > 0
                ? paragraph.FontSizePoints.Value
                : characterStyleDefaults.FontSize ?? styleDefaults.FontSize ?? tableRunStyleDefaults.FontSize;
            fontSize = ResolveNativeComplexScriptFontSize(fontSize, runProperties, characterStyleDefaults, styleDefaults, tableRunStyleDefaults, resolvedNativeDefaults);
            PdfCore.PdfStandardFont? font = ResolveNativeTextRunFont(paragraph, fallback, characterStyleDefaults, styleDefaults, tableRunStyleDefaults, resolvedNativeDefaults, nativeFontMap);
            string? fontFamily = ResolveNativeTextRunFontFamily(paragraph, fallback, characterStyleDefaults, styleDefaults, tableRunStyleDefaults, resolvedNativeDefaults, nativeFontMap);

            PdfCore.PdfColor? color = TryGetNativeRunColor(runProperties, out PdfCore.PdfColor? directColor)
                ? directColor
                : ParseNativeColor(characterStyleDefaults.ColorHex) ?? ParseNativeColor(styleDefaults.ColorHex) ?? tableRunStyleDefaults.Color ?? ParseNativeColor(tableRunStyleDefaults.ColorHex);
            PdfCore.PdfColor? background = TryGetNativeRunHighlight(runProperties, out PdfCore.PdfColor? directBackground)
                ? directBackground
                : MapNativeHighlight(characterStyleDefaults.Highlight) ?? MapNativeHighlight(styleDefaults.Highlight) ?? MapNativeHighlight(tableRunStyleDefaults.Highlight);

            return new NativeResolvedTextStyle(bold, underlineStyle, italic, strikeStyle, allCaps, baseline, fontSize, font, fontFamily, color, background) {
                TextSpacing = resolvedNativeDefaults.TextSpacing.Merge(tableRunStyleDefaults.TextSpacing)
                    .Merge(styleDefaults.TextSpacing).Merge(characterStyleDefaults.TextSpacing).Merge(runProperties).Resolve(),
                // Numbering uses the paragraph mark, independently of the first body run.
                ListMarkerTextSpacing = resolvedNativeDefaults.TextSpacing.Merge(tableRunStyleDefaults.TextSpacing)
                    .Merge(styleDefaults.TextSpacing).Merge(markerCharacterStyle.TextSpacing).Merge(markerProperties).Resolve()
            };
        }

        private static void ResetNativeTextStyle(PdfCore.PdfParagraphBuilder builder) {
            builder.HorizontalTextScaling(100D)
                .CharacterSpacing(0D)
                .Bold(false)
                .Italic(false)
                .Underline(false)
                .Strike(false)
                .Baseline(PdfCore.PdfTextBaseline.Normal)
                .ResetColor()
                .ResetFontSize()
                .ResetFont()
                .ResetBackgroundColor();
        }

    }
}
