using System.Globalization;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using WP14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
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
    public sealed class WordDrawingLayoutInfo {
        internal WordDrawingLayoutInfo(
            WordDrawingPlacementKind placement,
            string name,
            OfficeLength width,
            OfficeLength height,
            string? horizontalRelativeFrom,
            string? verticalRelativeFrom,
            OfficeLength? horizontalOffset,
            OfficeLength? verticalOffset,
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
            Width = width;
            Height = height;
            HorizontalRelativeFrom = horizontalRelativeFrom;
            VerticalRelativeFrom = verticalRelativeFrom;
            HorizontalOffset = horizontalOffset;
            VerticalOffset = verticalOffset;
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

        // Point projections are internal conversion/comparison plumbing, not public geometry.
        internal double WidthPoints => Width.ToPoints();
        internal double HeightPoints => Height.ToPoints();
        internal double? HorizontalOffsetPoints => HorizontalOffset?.ToPoints();
        internal double? VerticalOffsetPoints => VerticalOffset?.ToPoints();

        /// <summary>Gets the persisted placement kind.</summary>
        public WordDrawingPlacementKind Placement { get; }
        /// <summary>Gets the producer-supplied non-visual drawing name.</summary>
        public string Name { get; }
        /// <summary>Gets the persisted physical frame width.</summary>
        public OfficeLength Width { get; }
        /// <summary>Gets the persisted physical frame height.</summary>
        public OfficeLength Height { get; }
        /// <summary>Gets the horizontal reference token for anchored content.</summary>
        public string? HorizontalRelativeFrom { get; }
        /// <summary>Gets the vertical reference token for anchored content.</summary>
        public string? VerticalRelativeFrom { get; }
        /// <summary>Gets the explicit physical horizontal offset, when the producer used an offset.</summary>
        public OfficeLength? HorizontalOffset { get; }
        /// <summary>Gets the explicit physical vertical offset, when the producer used an offset.</summary>
        public OfficeLength? VerticalOffset { get; }
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
        internal static bool TryRead(WordDrawing? drawing, out WordDrawingLayoutInfo snapshot) =>
            TryRead(OfficeXmlNode.FromSdk(drawing), out snapshot);

        /// <summary>
        /// Reads the persisted frame of a <c>w:drawing</c>. Values are read as the SDK's typed members read them, so a
        /// malformed number throws <see cref="FormatException"/>.
        /// </summary>
        internal static bool TryRead(OfficeXmlNode drawing, out WordDrawingLayoutInfo snapshot) {
            OfficeXmlNode inline = drawing.IsNone ? default : drawing.Child(Wp, "inline");
            OfficeXmlNode inlineExtent = inline.IsNone ? default : inline.Child(Wp, "extent");
            if (ReadInt64(inlineExtent, "cx") is long inlineWidth && ReadInt64(inlineExtent, "cy") is long inlineHeight) {
                snapshot = new WordDrawingLayoutInfo(
                    WordDrawingPlacementKind.Inline,
                    ReadName(inline),
                    OfficeLength.Emus(inlineWidth),
                    OfficeLength.Emus(inlineHeight),
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

            OfficeXmlNode anchor = drawing.IsNone ? default : drawing.Child(Wp, "anchor");
            OfficeXmlNode anchorExtent = anchor.IsNone ? default : anchor.Child(Wp, "extent");
            if (ReadInt64(anchorExtent, "cx") is long anchorWidth && ReadInt64(anchorExtent, "cy") is long anchorHeight) {
                bool usesSimplePosition = ReadBoolean(anchor, "simplePos") == true;
                OfficeXmlNode horizontal = anchor.Child(Wp, "positionH");
                OfficeXmlNode vertical = anchor.Child(Wp, "positionV");
                OfficeXmlNode simplePosition = anchor.Child(Wp, "simplePos");
                OfficeLength? horizontalOffset = usesSimplePosition
                    ? ReadCoordinate(ReadInt64(simplePosition, "x"))
                    : ReadOffset(ChildText(horizontal, "posOffset"));
                OfficeLength? verticalOffset = usesSimplePosition
                    ? ReadCoordinate(ReadInt64(simplePosition, "y"))
                    : ReadOffset(ChildText(vertical, "posOffset"));
                snapshot = new WordDrawingLayoutInfo(
                    WordDrawingPlacementKind.Anchored,
                    ReadName(anchor),
                    OfficeLength.Emus(anchorWidth),
                    OfficeLength.Emus(anchorHeight),
                    usesSimplePosition ? "page" : ReadAttribute(horizontal, "relativeFrom"),
                    usesSimplePosition ? "page" : ReadAttribute(vertical, "relativeFrom"),
                    horizontalOffset,
                    verticalOffset,
                    usesSimplePosition ? null : ChildText(horizontal, "align"),
                    usesSimplePosition ? null : ChildText(vertical, "align"),
                    usesSimplePosition,
                    ReadWrap(anchor),
                    ReadBoolean(anchor, "behindDoc") ?? false,
                    ReadBoolean(anchor, "layoutInCell") ?? true,
                    ReadBoolean(anchor, "allowOverlap") ?? true,
                    ReadUInt32(anchor, "relativeHeight") ?? 0U,
                    IsGroupDrawing(drawing));
                return true;
            }

            snapshot = null!;
            return false;
        }

        internal static void SetExtent(WordDrawing drawing, long? widthEmu, long? heightEmu) {
            if (drawing.Inline?.Extent != null) {
                if (widthEmu.HasValue) drawing.Inline.Extent.Cx = widthEmu.Value;
                if (heightEmu.HasValue) drawing.Inline.Extent.Cy = heightEmu.Value;
                return;
            }

            if (drawing.Anchor?.Extent != null) {
                if (widthEmu.HasValue) drawing.Anchor.Extent.Cx = widthEmu.Value;
                if (heightEmu.HasValue) drawing.Anchor.Extent.Cy = heightEmu.Value;
                if (widthEmu.HasValue) drawing.Anchor.GetFirstChild<WP14.RelativeWidth>()?.Remove();
                if (heightEmu.HasValue) drawing.Anchor.GetFirstChild<WP14.RelativeHeight>()?.Remove();
            }
        }

        private const int Wp = OfficeXmlNamespaces.WordprocessingDrawing;

        /// <summary>
        /// Whether a graphic data element of the drawing declares a group payload. Graphic data inside an extension list
        /// doesn't count: the SDK read extension content as unknown elements, which its typed lookup skipped.
        /// </summary>
        private static bool IsGroupDrawing(OfficeXmlNode drawing) {
            foreach (OfficeXmlNode data in drawing.Descendants(OfficeXmlNamespaces.DrawingML, "graphicData")) {
                if (data.Attribute("uri")?.IndexOf("wordprocessingGroup", StringComparison.OrdinalIgnoreCase) >= 0 && !IsInsideExtensionList(data, drawing)) return true;
            }
            return false;
        }

        private static bool IsInsideExtensionList(OfficeXmlNode element, OfficeXmlNode drawing) {
            for (OfficeXmlNode ancestor = element.Parent; !ancestor.IsNone && ancestor != drawing; ancestor = ancestor.Parent) {
                if (string.Equals(ancestor.LocalName, "extLst", StringComparison.Ordinal)) return true;
            }
            return false;
        }

        private static string ReadName(OfficeXmlNode frame) {
            OfficeXmlNode properties = frame.Child(Wp, "docPr");
            return (properties.IsNone ? null : properties.Attribute("name")) ?? string.Empty;
        }

        private static string? ChildText(OfficeXmlNode parent, string localName) {
            if (parent.IsNone) return null;
            OfficeXmlNode child = parent.Child(Wp, localName);
            return child.IsNone ? null : child.InnerText;
        }

        private static long? ReadInt64(OfficeXmlNode element, string localName) {
            string? value = element.IsNone ? null : element.Attribute(localName);
            return value == null ? null : XmlConvert.ToInt64(value);
        }

        private static uint? ReadUInt32(OfficeXmlNode element, string localName) {
            string? value = element.Attribute(localName);
            return value == null ? null : XmlConvert.ToUInt32(value);
        }

        private static bool? ReadBoolean(OfficeXmlNode element, string localName) {
            string? value = element.Attribute(localName);
            return value == null ? null : XmlConvert.ToBoolean(value);
        }

        private static OfficeLength? ReadOffset(string? text) =>
            long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu)
                ? OfficeLength.Emus(emu)
                : null;

        private static OfficeLength? ReadCoordinate(long? emu) =>
            emu.HasValue ? OfficeLength.Emus(emu.Value) : null;

        // As the SDK's GetAttribute reads it: a missing attribute on a present element throws KeyNotFoundException.
        private static string? ReadAttribute(OfficeXmlNode element, string localName) {
            if (element.IsNone) return null;
            string value = element.Attribute(localName) ?? throw new KeyNotFoundException();
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        private static WordDrawingWrapKind ReadWrap(OfficeXmlNode anchor) {
            if (!anchor.Child(Wp, "wrapNone").IsNone) return WordDrawingWrapKind.None;
            if (!anchor.Child(Wp, "wrapSquare").IsNone) return WordDrawingWrapKind.Square;
            if (!anchor.Child(Wp, "wrapTight").IsNone) return WordDrawingWrapKind.Tight;
            if (!anchor.Child(Wp, "wrapThrough").IsNone) return WordDrawingWrapKind.Through;
            if (!anchor.Child(Wp, "wrapTopAndBottom").IsNone) return WordDrawingWrapKind.TopAndBottom;
            return WordDrawingWrapKind.Unsupported;
        }
    }
}
