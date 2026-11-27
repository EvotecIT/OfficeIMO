using System.Globalization;

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

        public NativeParagraphSpacing Merge(OfficeXmlNode properties) {
            OfficeXmlNode spacing = properties.WChild("spacing");
            return new NativeParagraphSpacing {
                BeforeLines = ReadLineUnits(spacing.WAttr("beforeLines")),
                AfterLines = ReadLineUnits(spacing.WAttr("afterLines")),
                BeforeAuto = WordXml.ParseOnOff(spacing.WAttr("beforeAutospacing")),
                AfterAuto = WordXml.ParseOnOff(spacing.WAttr("afterAutospacing")),
                SnapToGrid = ReadNativeOnOff(properties.WChild("snapToGrid"))
            }.Inherit(this);
        }

        private static int? ReadLineUnits(string? value) =>
            int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int units) && units >= 0 ? units : null;
    }

    private static NativeParagraphSpacing ResolveNativeParagraphSpacing(WordParagraph paragraph,
        NativeParagraphStyleDefaults styles, NativeDocumentDefaults defaults, NativeTableStyleDefaults? table = null) =>
        styles.ParagraphSpacing.Inherit((table ?? NativeTableStyleDefaults.Empty).ParagraphSpacing.Inherit(defaults.ParagraphSpacing))
            .Merge(paragraph.Node().WChild("pPr"));

    private static (double? Before, double? After) ResolveNativeParagraphSpacingInLines(WordParagraph paragraph,
        NativeParagraphStyleDefaults styles, NativeDocumentDefaults defaults, NativeTableStyleDefaults? table = null) {
        NativeParagraphSpacing spacing = ResolveNativeParagraphSpacing(paragraph, styles, defaults, table);
        // Word uses the section's line-grid pitch even when this paragraph
        // disables snapping. Without an active grid one line unit is 12pt.
        double unit = GetNativeSectionLineGridPitch(paragraph, defaults.NoteContentSection) ?? 12D;
        return (spacing.BeforeAuto != true && spacing.BeforeLines.HasValue ? spacing.BeforeLines.Value / 100D * unit : null,
            spacing.AfterAuto != true && spacing.AfterLines.HasValue ? spacing.AfterLines.Value / 100D * unit : null);
    }

    private static double? GetNativeSectionLineGridPitch(WordParagraph paragraph, WordSection? sourceSection = null) {
        OfficeXmlNode grid = (sourceSection ?? ResolveNativeFrameSourceParagraph(paragraph).GetBodySection())?.Node().WChild("docGrid") ?? default;
        return grid.WAttr("type") is { } type && type != "default" &&
            grid.WInt32("linePitch") is int pitch && pitch > 0 ? pitch / 20D : null;
    }

    private static bool RenderNativeDocumentStartSpacing(INativePdfFlow pdf, WordParagraph paragraph, double spacingBefore) {
        if (spacingBefore <= 0D || paragraph.Node() != paragraph._document.BodyNode().FirstDescendant(WordXml.W, "p")) return false;
        // Word spacing values can combine with the section line grid into an
        // effectively unbounded number of empty column frames. Reject such
        // geometry before the PDF writer begins advancing through them.
        if (double.IsNaN(spacingBefore) || double.IsInfinity(spacingBefore) || spacingBefore > 1_000_000D)
            throw new InvalidDataException("Initial paragraph spacing exceeds the supported PDF layout range.");
        // Initial document spacing is authored content. Reserve it explicitly
        // so normal page/column breaks continue to suppress spacing at the top.
        pdf.Spacer(spacingBefore);
        return true;
    }
}
