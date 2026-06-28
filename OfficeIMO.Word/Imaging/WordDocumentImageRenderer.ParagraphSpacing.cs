using System;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private readonly struct WordParagraphSpacing {
            public WordParagraphSpacing(double before, double after) {
                Before = before;
                After = after;
            }

            public double Before { get; }

            public double After { get; }
        }

        private readonly struct WordParagraphSpacingState {
            public WordParagraphSpacingState(string styleKey, double after, bool contextualSpacing) {
                StyleKey = styleKey;
                After = after;
                ContextualSpacing = contextualSpacing;
            }

            public string StyleKey { get; }

            public double After { get; }

            public bool ContextualSpacing { get; }
        }

        private static WordParagraphSpacing ResolveParagraphSpacing(WordParagraph paragraph, double fontSize, double lineHeight, WordImageFlowContext context, out WordParagraphSpacingState state) {
            double before = ResolveParagraphSpacingBefore(paragraph, fontSize, lineHeight) ?? 0D;
            double after = ResolveParagraphSpacingAfter(paragraph, fontSize, lineHeight) ?? ParagraphGapPoints;
            bool contextualSpacing = ResolveParagraphContextualSpacing(paragraph);
            string styleKey = ResolveParagraphStyleKey(paragraph);
            WordParagraphSpacing spacing = new WordParagraphSpacing(Math.Max(0D, before), Math.Max(0D, after));
            WordParagraphSpacingState? previousState = context.PreviousParagraphSpacingState;
            if (previousState.HasValue &&
                string.Equals(previousState.Value.StyleKey, styleKey, StringComparison.OrdinalIgnoreCase) &&
                (previousState.Value.ContextualSpacing || contextualSpacing)) {
                context.Y = Math.Max(0D, context.Y - previousState.Value.After);
                spacing = new WordParagraphSpacing(0D, spacing.After);
            }

            state = new WordParagraphSpacingState(styleKey, spacing.After, contextualSpacing);
            return spacing;
        }

        private static double? ResolveParagraphSpacingBefore(WordParagraph paragraph, double fontSize, double lineHeight) {
            SpacingBetweenLines? directSpacing = paragraph._paragraphProperties?.GetFirstChild<SpacingBetweenLines>();
            if (HasSpacingBefore(directSpacing)) {
                return ResolveSpacingBeforePoints(directSpacing, fontSize, lineHeight);
            }

            foreach (StyleParagraphProperties properties in EnumerateParagraphStyleProperties(paragraph)) {
                SpacingBetweenLines? styleSpacing = properties.GetFirstChild<SpacingBetweenLines>();
                if (HasSpacingBefore(styleSpacing)) {
                    return ResolveSpacingBeforePoints(styleSpacing, fontSize, lineHeight);
                }
            }

            SpacingBetweenLines? defaultSpacing = GetDocumentDefaultParagraphSpacing(paragraph);
            if (HasSpacingBefore(defaultSpacing)) {
                return ResolveSpacingBeforePoints(defaultSpacing, fontSize, lineHeight);
            }

            return null;
        }

        private static double? ResolveParagraphSpacingAfter(WordParagraph paragraph, double fontSize, double lineHeight) {
            SpacingBetweenLines? directSpacing = paragraph._paragraphProperties?.GetFirstChild<SpacingBetweenLines>();
            if (HasSpacingAfter(directSpacing)) {
                return ResolveSpacingAfterPoints(directSpacing, fontSize, lineHeight);
            }

            foreach (StyleParagraphProperties properties in EnumerateParagraphStyleProperties(paragraph)) {
                SpacingBetweenLines? styleSpacing = properties.GetFirstChild<SpacingBetweenLines>();
                if (HasSpacingAfter(styleSpacing)) {
                    return ResolveSpacingAfterPoints(styleSpacing, fontSize, lineHeight);
                }
            }

            SpacingBetweenLines? defaultSpacing = GetDocumentDefaultParagraphSpacing(paragraph);
            if (HasSpacingAfter(defaultSpacing)) {
                return ResolveSpacingAfterPoints(defaultSpacing, fontSize, lineHeight);
            }

            return null;
        }

        private static SpacingBetweenLines? GetDocumentDefaultParagraphSpacing(WordParagraph paragraph) =>
            paragraph._document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?
                .GetFirstChild<ParagraphPropertiesDefault>()?
                .GetFirstChild<ParagraphPropertiesBaseStyle>()?
                .GetFirstChild<SpacingBetweenLines>();

        private static bool ResolveParagraphContextualSpacing(WordParagraph paragraph) {
            ContextualSpacing? directSpacing = paragraph._paragraphProperties?.GetFirstChild<ContextualSpacing>();
            if (directSpacing != null) {
                return IsContextualSpacingEnabled(directSpacing);
            }

            foreach (StyleParagraphProperties properties in EnumerateParagraphStyleProperties(paragraph)) {
                ContextualSpacing? styleSpacing = properties.GetFirstChild<ContextualSpacing>();
                if (styleSpacing != null) {
                    return IsContextualSpacingEnabled(styleSpacing);
                }
            }

            return false;
        }

        private static bool IsContextualSpacingEnabled(ContextualSpacing spacing) =>
            spacing.Val?.Value != false;

        private static string ResolveParagraphStyleKey(WordParagraph paragraph) =>
            string.IsNullOrWhiteSpace(paragraph.StyleId) ? string.Empty : paragraph.StyleId!;

        private static bool HasSpacingBefore(SpacingBetweenLines? spacing) =>
            spacing?.Before != null || spacing?.BeforeLines != null || spacing?.BeforeAutoSpacing != null;

        private static bool HasSpacingAfter(SpacingBetweenLines? spacing) =>
            spacing?.After != null || spacing?.AfterLines != null || spacing?.AfterAutoSpacing != null;

        private static double? ResolveSpacingBeforePoints(SpacingBetweenLines? spacing, double fontSize, double lineHeight) {
            if (spacing == null || IsOnOffTrue(spacing.BeforeAutoSpacing)) {
                return null;
            }

            return ConvertLineHundredthsToPoints(spacing.BeforeLines, fontSize, lineHeight)
                ?? ConvertTwipsToPoints(spacing.Before?.Value);
        }

        private static double? ResolveSpacingAfterPoints(SpacingBetweenLines? spacing, double fontSize, double lineHeight) {
            if (spacing == null || IsOnOffTrue(spacing.AfterAutoSpacing)) {
                return null;
            }

            return ConvertLineHundredthsToPoints(spacing.AfterLines, fontSize, lineHeight)
                ?? ConvertTwipsToPoints(spacing.After?.Value);
        }

        private static double? ConvertLineHundredthsToPoints(Int32Value? value, double fontSize, double lineHeight) {
            if (value == null ||
                value.Value < 0 ||
                lineHeight <= 0D ||
                double.IsNaN(lineHeight) ||
                double.IsInfinity(lineHeight)) {
                return null;
            }

            return (value.Value / 100D) * lineHeight;
        }

        private static double? ConvertTwipsToPoints(string? value) {
            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double twips) ||
                twips < 0D ||
                double.IsNaN(twips) ||
                double.IsInfinity(twips)) {
                return null;
            }

            return twips / TwipsPerPoint;
        }

        private static bool IsOnOffTrue(OnOffValue? value) =>
            value?.Value == true;
    }
}
