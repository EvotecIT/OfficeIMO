using System.Globalization;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private static void ApplyWordParagraphFormatting(WordParagraphSnapshot source, OdtParagraph target) {
        if (source.Alignment != null && TryMapWordAlignment(source.Alignment, out OdtParagraphAlignment alignment)) {
            target.Alignment = alignment;
        }
        if (source.IndentStartPoints.HasValue) target.IndentStart = OdfLength.Points(source.IndentStartPoints.Value);
        if (source.IndentEndPoints.HasValue) target.IndentEnd = OdfLength.Points(source.IndentEndPoints.Value);
        if (source.IndentFirstLinePoints.HasValue) target.FirstLineIndent = OdfLength.Points(source.IndentFirstLinePoints.Value);
        if (source.SpaceAbovePoints.HasValue) target.SpaceAbove = OdfLength.Points(source.SpaceAbovePoints.Value);
        if (source.SpaceBelowPoints.HasValue) target.SpaceBelow = OdfLength.Points(source.SpaceBelowPoints.Value);
        if (source.IsRightToLeft) target.WritingMode = "rl-tb";
        if (TryMapWordLineHeight(source, out OdfLength lineHeight)) target.LineHeight = lineHeight;
        if (OdfColor.TryParse(source.ShadingFillColorHex, out OdfColor background)) target.BackgroundColor = background;
    }

    private static int ApplyOdtParagraphFormatting(OdtParagraph source, WordParagraph target) {
        int unsupported = 0;
        switch (source.Alignment) {
            case OdtParagraphAlignment.Start: target.ParagraphAlignment = WordParagraphAlignment.Start; break;
            case OdtParagraphAlignment.Left: target.ParagraphAlignment = WordParagraphAlignment.Left; break;
            case OdtParagraphAlignment.Center: target.ParagraphAlignment = WordParagraphAlignment.Center; break;
            case OdtParagraphAlignment.Right: target.ParagraphAlignment = WordParagraphAlignment.Right; break;
            case OdtParagraphAlignment.End: target.ParagraphAlignment = WordParagraphAlignment.End; break;
            case OdtParagraphAlignment.Justify: target.ParagraphAlignment = WordParagraphAlignment.Both; break;
        }
        target.BiDi = source.IsRightToLeft;
        if (source.IndentStart.HasValue) {
            if (source.IndentStart.Value.TryToPoints(out double points)) target.IndentationBeforePoints = points; else unsupported++;
        }
        if (source.IndentEnd.HasValue) {
            if (source.IndentEnd.Value.TryToPoints(out double points)) target.IndentationAfterPoints = points; else unsupported++;
        }
        if (source.FirstLineIndent.HasValue) {
            if (source.FirstLineIndent.Value.TryToPoints(out double points)) target.IndentationFirstLinePoints = points; else unsupported++;
        }
        if (source.SpaceAbove.HasValue) {
            if (source.SpaceAbove.Value.TryToPoints(out double points)) target.LineSpacingBeforePoints = points; else unsupported++;
        }
        if (source.SpaceBelow.HasValue) {
            if (source.SpaceBelow.Value.TryToPoints(out double points)) target.LineSpacingAfterPoints = points; else unsupported++;
        }
        unsupported += ApplyOdtLineHeight(source.LineHeight, target);
        if (source.BackgroundColor.HasValue) target.ShadingFillColorHex = source.BackgroundColor.Value.ToString();
        return unsupported;
    }

    private static int ApplyOdtLineHeight(OdfLength? lineHeight, WordParagraph target) {
        if (!lineHeight.HasValue) return 0;
        string lexical = lineHeight.Value.ToString();
        if (string.Equals(lexical, "normal", StringComparison.OrdinalIgnoreCase)) return 0;
        if (lexical.EndsWith("%", StringComparison.Ordinal)
            && double.TryParse(lexical.Substring(0, lexical.Length - 1), NumberStyles.Float,
                CultureInfo.InvariantCulture, out double percentage)
            && percentage > 0D && percentage <= int.MaxValue * 100D / 240D) {
            target.LineSpacingRule = WordLineSpacingRule.Auto;
            target.LineSpacing = checked((int)Math.Round(percentage * 240D / 100D, MidpointRounding.AwayFromZero));
            return 0;
        }
        if (lineHeight.Value.TryToPoints(out double points) && points >= 0D) {
            target.LineSpacingRule = WordLineSpacingRule.Exact;
            target.LineSpacingPoints = points;
            return 0;
        }
        return 1;
    }

    private static bool TryMapWordLineHeight(WordParagraphSnapshot source, out OdfLength lineHeight) {
        lineHeight = default;
        if (!source.LineSpacingValue.HasValue || source.LineSpacingValue.Value < 0) return false;
        if (source.LineSpacingRule == null
            || string.Equals(source.LineSpacingRule, "auto", StringComparison.OrdinalIgnoreCase)) {
            double percentage = source.LineSpacingValue.Value * 100D / 240D;
            lineHeight = OdfLength.Parse(percentage.ToString("0.###", CultureInfo.InvariantCulture) + "%");
            return true;
        }
        if (string.Equals(source.LineSpacingRule, "exact", StringComparison.OrdinalIgnoreCase)) {
            lineHeight = OdfLength.Points(source.LineSpacingValue.Value / 20D);
            return true;
        }
        return false;
    }
}
