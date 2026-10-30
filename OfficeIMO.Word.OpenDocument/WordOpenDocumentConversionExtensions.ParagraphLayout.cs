using System.Globalization;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private static void ApplyWordParagraphFormatting(WordParagraphInfo source, OdtParagraph target) {
        if (source.Alignment != null && TryMapWordAlignment(source.Alignment, out OdtParagraphAlignment alignment)) {
            target.Alignment = alignment;
        }
        if (source.IndentStart.HasValue) target.IndentStart = OdfLength.Points(source.IndentStart.Value.ToPoints());
        if (source.IndentEnd.HasValue) target.IndentEnd = OdfLength.Points(source.IndentEnd.Value.ToPoints());
        if (source.IndentFirstLine.HasValue) target.FirstLineIndent = OdfLength.Points(source.IndentFirstLine.Value.ToPoints());
        if (source.SpaceAbove.HasValue) target.SpaceAbove = OdfLength.Points(source.SpaceAbove.Value.ToPoints());
        if (source.SpaceBelow.HasValue) target.SpaceBelow = OdfLength.Points(source.SpaceBelow.Value.ToPoints());
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
            if (source.IndentStart.Value.TryToPoints(out double points)) target.IndentationBefore = global::OfficeIMO.OfficeLength.Points(points); else unsupported++;
        }
        if (source.IndentEnd.HasValue) {
            if (source.IndentEnd.Value.TryToPoints(out double points)) target.IndentationAfter = global::OfficeIMO.OfficeLength.Points(points); else unsupported++;
        }
        if (source.FirstLineIndent.HasValue) {
            if (source.FirstLineIndent.Value.TryToPoints(out double points)) target.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(points); else unsupported++;
        }
        if (source.SpaceAbove.HasValue) {
            if (source.SpaceAbove.Value.TryToPoints(out double points)) target.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(points); else unsupported++;
        }
        if (source.SpaceBelow.HasValue) {
            if (source.SpaceBelow.Value.TryToPoints(out double points)) target.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(points); else unsupported++;
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
            target.LineSpacing = WordLineSpacing.Multiple(percentage / 100D);
            return 0;
        }
        if (lineHeight.Value.TryToPoints(out double points) && points >= 0D) {
            target.LineSpacing = WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(points));
            return 0;
        }
        return 1;
    }

    private static bool TryMapWordLineHeight(WordParagraphInfo source, out OdfLength lineHeight) {
        lineHeight = default;
        if (source.LineSpacing is not WordLineSpacing spacing) return false;
        if (spacing.MultipleValue is double multiple) {
            lineHeight = OdfLength.Parse((multiple * 100D).ToString("0.###", CultureInfo.InvariantCulture) + "%");
            return true;
        }
        if (spacing.Rule == WordLineSpacingRule.Exact && spacing.Height.HasValue) {
            lineHeight = OdfLength.Points(spacing.Height.Value.ToPoints());
            return true;
        }
        return false;
    }
}
