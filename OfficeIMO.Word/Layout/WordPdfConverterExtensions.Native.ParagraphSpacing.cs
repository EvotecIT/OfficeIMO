using DocumentFormat.OpenXml;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // Line units and grid snapping inherit independently of the point-valued
    // spacing attributes. Keep them authored until the section is known.
    private readonly record struct NativeParagraphSpacing {
        public int? BeforeLines { get; init; }
        public int? AfterLines { get; init; }
        public bool? BeforeAuto { get; init; }
        public bool? AfterAuto { get; init; }
        public bool? SnapToGrid { get; init; }

        public NativeParagraphSpacing Inherit(NativeParagraphSpacing inherited) => new() {
            BeforeLines = BeforeLines ?? inherited.BeforeLines,
            AfterLines = AfterLines ?? inherited.AfterLines,
            BeforeAuto = BeforeAuto ?? inherited.BeforeAuto,
            AfterAuto = AfterAuto ?? inherited.AfterAuto,
            SnapToGrid = SnapToGrid ?? inherited.SnapToGrid
        };

        public NativeParagraphSpacing Merge(OpenXmlElement? properties) {
            W.SpacingBetweenLines? spacing = properties?.GetFirstChild<W.SpacingBetweenLines>();
            return new NativeParagraphSpacing {
                BeforeLines = ReadLineUnits(spacing?.BeforeLines),
                AfterLines = ReadLineUnits(spacing?.AfterLines),
                BeforeAuto = spacing?.BeforeAutoSpacing?.Value,
                AfterAuto = spacing?.AfterAutoSpacing?.Value,
                SnapToGrid = ReadNativeOnOff(properties?.GetFirstChild<W.SnapToGrid>())
            }.Inherit(this);
        }

        private static int? ReadLineUnits(Int32Value? value) =>
            int.TryParse(value?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int units) && units >= 0 ? units : null;
    }

    private static NativeParagraphSpacing ResolveNativeParagraphSpacing(WordParagraph paragraph,
        NativeParagraphStyleDefaults styles, NativeDocumentDefaults defaults, NativeTableStyleDefaults table = default) =>
        styles.ParagraphSpacing.Inherit(table.ParagraphSpacing.Inherit(defaults.ParagraphSpacing))
            .Merge(paragraph._paragraph.ParagraphProperties);

    private static (double? Before, double? After) ResolveNativeParagraphSpacingInLines(WordParagraph paragraph,
        NativeParagraphStyleDefaults styles, NativeDocumentDefaults defaults, NativeTableStyleDefaults table = default) {
        NativeParagraphSpacing spacing = ResolveNativeParagraphSpacing(paragraph, styles, defaults, table);
        // Word uses the section's line-grid pitch even when this paragraph
        // disables snapping. Without an active grid one line unit is 12pt.
        double unit = GetNativeSectionLineGridPitch(paragraph) ?? 12D;
        return (spacing.BeforeAuto != true && spacing.BeforeLines.HasValue ? spacing.BeforeLines.Value / 100D * unit : null,
            spacing.AfterAuto != true && spacing.AfterLines.HasValue ? spacing.AfterLines.Value / 100D * unit : null);
    }

    private static double? GetNativeSectionLineGridPitch(WordParagraph paragraph) {
        W.DocGrid? grid = paragraph.GetBodySection()?._sectionProperties.GetFirstChild<W.DocGrid>();
        return grid?.Type?.Value != null && grid.Type.Value != W.DocGridValues.Default &&
            grid.LinePitch?.Value is int pitch && pitch > 0 ? pitch / 20D : null;
    }

    private static bool RenderNativeDocumentStartSpacing(INativePdfFlow pdf, WordParagraph paragraph, double spacingBefore) {
        if (spacingBefore <= 0D || !ReferenceEquals(paragraph._paragraph, paragraph._document
            ._wordprocessingDocument.MainDocumentPart?.Document?.Body?.Descendants<W.Paragraph>().FirstOrDefault())) return false;
        // Initial document spacing is authored content. Reserve it explicitly
        // so normal page/column breaks continue to suppress spacing at the top.
        pdf.Spacer(spacingBefore);
        return true;
    }
}
