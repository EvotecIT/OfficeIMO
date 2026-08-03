using System.Globalization;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;

namespace OfficeIMO.Word {
    /// <summary>Identifies how DrawingML content is persisted in the Word package.</summary>
    public enum WordDrawingPlacementKind {
        /// <summary>The drawing participates in the text line.</summary>
        Inline,
        /// <summary>The drawing uses an anchored floating frame.</summary>
        Anchored
    }

    /// <summary>Identifies the persisted wrapping primitive on an anchored drawing.</summary>
    public enum WordDrawingWrapKind {
        /// <summary>Inline content has no floating wrap primitive.</summary>
        Inline,
        /// <summary>No text wrapping.</summary>
        None,
        /// <summary>Square wrapping.</summary>
        Square,
        /// <summary>Tight wrapping.</summary>
        Tight,
        /// <summary>Through wrapping.</summary>
        Through,
        /// <summary>Top-and-bottom wrapping.</summary>
        TopAndBottom,
        /// <summary>The producer emitted a wrapping shape OfficeIMO does not map.</summary>
        Unsupported
    }

    /// <summary>
    /// Captures persisted DrawingML geometry and anchoring evidence. Values describe package markup,
    /// not a claim about Word's final rendered coordinates after pagination and compatibility layout.
    /// </summary>
    public sealed class WordDrawingLayoutSnapshot {
        internal WordDrawingLayoutSnapshot(
            WordDrawingPlacementKind placement,
            string name,
            double widthPoints,
            double heightPoints,
            string? horizontalRelativeFrom,
            string? verticalRelativeFrom,
            double? horizontalOffsetPoints,
            double? verticalOffsetPoints,
            string? horizontalAlignment,
            string? verticalAlignment,
            bool usesSimplePosition,
            WordDrawingWrapKind wrap,
            bool behindDocument,
            bool layoutInCell,
            bool allowOverlap,
            uint relativeHeight,
            bool isGroup) {
            Placement = placement;
            Name = name;
            WidthPoints = widthPoints;
            HeightPoints = heightPoints;
            HorizontalRelativeFrom = horizontalRelativeFrom;
            VerticalRelativeFrom = verticalRelativeFrom;
            HorizontalOffsetPoints = horizontalOffsetPoints;
            VerticalOffsetPoints = verticalOffsetPoints;
            HorizontalAlignment = horizontalAlignment;
            VerticalAlignment = verticalAlignment;
            UsesSimplePosition = usesSimplePosition;
            Wrap = wrap;
            BehindDocument = behindDocument;
            LayoutInCell = layoutInCell;
            AllowOverlap = allowOverlap;
            RelativeHeight = relativeHeight;
            IsGroup = isGroup;
        }

        /// <summary>Gets the persisted placement kind.</summary>
        public WordDrawingPlacementKind Placement { get; }
        /// <summary>Gets the producer-supplied non-visual drawing name.</summary>
        public string Name { get; }
        /// <summary>Gets the persisted frame width in points.</summary>
        public double WidthPoints { get; }
        /// <summary>Gets the persisted frame height in points.</summary>
        public double HeightPoints { get; }
        /// <summary>Gets the horizontal reference token for anchored content.</summary>
        public string? HorizontalRelativeFrom { get; }
        /// <summary>Gets the vertical reference token for anchored content.</summary>
        public string? VerticalRelativeFrom { get; }
        /// <summary>Gets the explicit horizontal offset in points, when the producer used an offset.</summary>
        public double? HorizontalOffsetPoints { get; }
        /// <summary>Gets the explicit vertical offset in points, when the producer used an offset.</summary>
        public double? VerticalOffsetPoints { get; }
        /// <summary>Gets the horizontal alignment token when the producer used alignment instead of an offset.</summary>
        public string? HorizontalAlignment { get; }
        /// <summary>Gets the vertical alignment token when the producer used alignment instead of an offset.</summary>
        public string? VerticalAlignment { get; }
        /// <summary>Gets whether an anchored producer selected page-relative simple positioning.</summary>
        public bool UsesSimplePosition { get; }
        /// <summary>Gets the persisted wrap primitive.</summary>
        public WordDrawingWrapKind Wrap { get; }
        /// <summary>Gets whether the anchor is behind document text.</summary>
        public bool BehindDocument { get; }
        /// <summary>Gets whether the anchor participates in table-cell layout.</summary>
        public bool LayoutInCell { get; }
        /// <summary>Gets whether the anchor permits overlap with other floating objects.</summary>
        public bool AllowOverlap { get; }
        /// <summary>Gets the persisted relative z-order value.</summary>
        public uint RelativeHeight { get; }
        /// <summary>Gets whether the graphic data declares a WordprocessingGroup payload.</summary>
        public bool IsGroup { get; }
    }

    internal static class WordDrawingLayoutReader {
        private const double EmusPerPoint = 12700D;

        internal static bool TryRead(WordDrawing? drawing, out WordDrawingLayoutSnapshot snapshot) {
            if (drawing?.Inline is DW.Inline inline && inline.Extent?.Cx?.Value is long inlineWidth && inline.Extent.Cy?.Value is long inlineHeight) {
                snapshot = new WordDrawingLayoutSnapshot(
                    WordDrawingPlacementKind.Inline,
                    inline.DocProperties?.Name?.Value ?? string.Empty,
                    inlineWidth / EmusPerPoint,
                    inlineHeight / EmusPerPoint,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    false,
                    WordDrawingWrapKind.Inline,
                    false,
                    true,
                    false,
                    0,
                    IsGroupDrawing(drawing));
                return true;
            }

            if (drawing?.Anchor is DW.Anchor anchor && anchor.Extent?.Cx?.Value is long anchorWidth && anchor.Extent.Cy?.Value is long anchorHeight) {
                bool usesSimplePosition = anchor.SimplePos?.Value == true;
                double? horizontalOffset = usesSimplePosition
                    ? ReadCoordinatePoints(anchor.SimplePosition?.X?.Value)
                    : ReadOffsetPoints(anchor.HorizontalPosition?.PositionOffset?.Text);
                double? verticalOffset = usesSimplePosition
                    ? ReadCoordinatePoints(anchor.SimplePosition?.Y?.Value)
                    : ReadOffsetPoints(anchor.VerticalPosition?.PositionOffset?.Text);
                snapshot = new WordDrawingLayoutSnapshot(
                    WordDrawingPlacementKind.Anchored,
                    anchor.GetFirstChild<DW.DocProperties>()?.Name?.Value ?? string.Empty,
                    anchorWidth / EmusPerPoint,
                    anchorHeight / EmusPerPoint,
                    usesSimplePosition ? "page" : ReadAttribute(anchor.HorizontalPosition, "relativeFrom"),
                    usesSimplePosition ? "page" : ReadAttribute(anchor.VerticalPosition, "relativeFrom"),
                    horizontalOffset,
                    verticalOffset,
                    usesSimplePosition ? null : anchor.HorizontalPosition?.HorizontalAlignment?.Text,
                    usesSimplePosition ? null : anchor.VerticalPosition?.VerticalAlignment?.Text,
                    usesSimplePosition,
                    ReadWrap(anchor),
                    anchor.BehindDoc?.Value ?? false,
                    anchor.LayoutInCell?.Value ?? true,
                    anchor.AllowOverlap?.Value ?? true,
                    anchor.RelativeHeight?.Value ?? 0U,
                    IsGroupDrawing(drawing));
                return true;
            }

            snapshot = null!;
            return false;
        }

        internal static void SetExtent(WordDrawing drawing, long widthEmu, long? heightEmu) {
            if (drawing.Inline?.Extent != null) {
                drawing.Inline.Extent.Cx = widthEmu;
                if (heightEmu.HasValue) drawing.Inline.Extent.Cy = heightEmu.Value;
                return;
            }

            if (drawing.Anchor?.Extent != null) {
                drawing.Anchor.Extent.Cx = widthEmu;
                if (heightEmu.HasValue) drawing.Anchor.Extent.Cy = heightEmu.Value;
            }
        }

        private static bool IsGroupDrawing(WordDrawing drawing) => drawing
            .Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>()
            .Any(data => data.Uri?.Value?.IndexOf("wordprocessingGroup", StringComparison.OrdinalIgnoreCase) >= 0);

        private static double? ReadOffsetPoints(string? text) =>
            long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu)
                ? emu / EmusPerPoint
                : null;

        private static double? ReadCoordinatePoints(long? emu) =>
            emu.HasValue ? emu.Value / EmusPerPoint : null;

        private static string? ReadAttribute(DocumentFormat.OpenXml.OpenXmlElement? element, string localName) {
            if (element == null) return null;
            string? value = element.GetAttribute(localName, string.Empty).Value;
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        private static WordDrawingWrapKind ReadWrap(DW.Anchor anchor) {
            if (anchor.GetFirstChild<DW.WrapNone>() != null) return WordDrawingWrapKind.None;
            if (anchor.GetFirstChild<DW.WrapSquare>() != null) return WordDrawingWrapKind.Square;
            if (anchor.GetFirstChild<DW.WrapTight>() != null) return WordDrawingWrapKind.Tight;
            if (anchor.GetFirstChild<DW.WrapThrough>() != null) return WordDrawingWrapKind.Through;
            if (anchor.GetFirstChild<DW.WrapTopBottom>() != null) return WordDrawingWrapKind.TopAndBottom;
            return WordDrawingWrapKind.Unsupported;
        }
    }
}
