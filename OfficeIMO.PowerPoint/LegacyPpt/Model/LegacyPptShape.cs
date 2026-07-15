using OfficeIMO.Drawing.Binary;
using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Identifies the shape kinds currently projected from binary PowerPoint files.</summary>
    public enum LegacyPptShapeKind {
        /// <summary>A text-bearing shape.</summary>
        TextBox,

        /// <summary>A rectangle auto shape.</summary>
        Rectangle,

        /// <summary>An ellipse auto shape.</summary>
        Ellipse,

        /// <summary>A line auto shape.</summary>
        Line,

        /// <summary>An OfficeArt AutoShape with a DrawingML preset-geometry equivalent.</summary>
        AutoShape,

        /// <summary>An OfficeArt connector with a DrawingML connector-geometry equivalent.</summary>
        Connector,

        /// <summary>A nested OfficeArt shape group with its own child coordinate system.</summary>
        Group,

        /// <summary>An OfficeArt picture frame with importable image data.</summary>
        Picture,

        /// <summary>A recognized OfficeArt shape that is preserved only as metadata.</summary>
        Unsupported
    }

    /// <summary>Identifies the relevant legacy placeholder types.</summary>
    public enum LegacyPptPlaceholderKind : byte {
        /// <summary>The shape is not a placeholder.</summary>
        None = 0x00,
        /// <summary>Master title placeholder.</summary>
        MasterTitle = 0x01,
        /// <summary>Master body placeholder.</summary>
        MasterBody = 0x02,
        /// <summary>Master centered title placeholder.</summary>
        MasterCenterTitle = 0x03,
        /// <summary>Master subtitle placeholder.</summary>
        MasterSubtitle = 0x04,
        /// <summary>Master notes slide image placeholder.</summary>
        MasterNotesSlideImage = 0x05,
        /// <summary>Master notes body placeholder.</summary>
        MasterNotesBody = 0x06,
        /// <summary>Master date placeholder.</summary>
        MasterDate = 0x07,
        /// <summary>Master slide number placeholder.</summary>
        MasterSlideNumber = 0x08,
        /// <summary>Master footer placeholder.</summary>
        MasterFooter = 0x09,
        /// <summary>Master header placeholder.</summary>
        MasterHeader = 0x0A,
        /// <summary>Notes slide image placeholder.</summary>
        NotesSlideImage = 0x0B,
        /// <summary>Notes body placeholder.</summary>
        NotesBody = 0x0C,
        /// <summary>Title placeholder.</summary>
        Title = 0x0D,
        /// <summary>Body placeholder.</summary>
        Body = 0x0E,
        /// <summary>Centered title placeholder.</summary>
        CenterTitle = 0x0F,
        /// <summary>Subtitle placeholder.</summary>
        Subtitle = 0x10,
        /// <summary>Vertical title placeholder.</summary>
        VerticalTitle = 0x11,
        /// <summary>Vertical body placeholder.</summary>
        VerticalBody = 0x12,
        /// <summary>Generic object placeholder.</summary>
        Object = 0x13,
        /// <summary>Chart placeholder.</summary>
        Graph = 0x14,
        /// <summary>Table placeholder.</summary>
        Table = 0x15,
        /// <summary>Clip-art placeholder.</summary>
        ClipArt = 0x16,
        /// <summary>Organization-chart placeholder.</summary>
        OrganizationChart = 0x17,
        /// <summary>Media placeholder.</summary>
        Media = 0x18,
        /// <summary>Vertical object placeholder.</summary>
        VerticalObject = 0x19,
        /// <summary>Picture placeholder.</summary>
        Picture = 0x1A
    }

    /// <summary>Represents a shape decoded from a binary PowerPoint slide.</summary>
    public sealed class LegacyPptShape {
        internal LegacyPptShape(LegacyPptShapeKind kind, ushort officeArtShapeType, uint shapeId,
            long recordOffset, LegacyPptBounds bounds, string text, LegacyPptPlaceholderKind placeholderKind,
            OfficeArtShapeStyle style, string? fillColor, string? lineColor,
            int? pictureStoreIndex = null, OfficeArtBlipStoreEntry? picture = null,
            LegacyPptBounds? groupCoordinateBounds = null,
            IReadOnlyList<LegacyPptShape>? children = null) {
            Kind = kind;
            OfficeArtShapeType = officeArtShapeType;
            ShapeId = shapeId;
            RecordOffset = recordOffset;
            Bounds = bounds;
            Text = text ?? string.Empty;
            PlaceholderKind = placeholderKind;
            Style = style ?? throw new ArgumentNullException(nameof(style));
            FillColor = fillColor;
            LineColor = lineColor;
            PictureStoreIndex = pictureStoreIndex;
            Picture = picture;
            GroupCoordinateBounds = groupCoordinateBounds;
            Children = new ReadOnlyCollection<LegacyPptShape>(
                children?.ToArray() ?? Array.Empty<LegacyPptShape>());
        }

        /// <summary>Gets the projected shape kind.</summary>
        public LegacyPptShapeKind Kind { get; }

        /// <summary>Gets the raw OfficeArt shape type.</summary>
        public ushort OfficeArtShapeType { get; }

        /// <summary>Gets the OfficeArt shape identifier.</summary>
        public uint ShapeId { get; }

        /// <summary>Gets the shape-container offset in the PowerPoint Document stream.</summary>
        public long RecordOffset { get; }

        /// <summary>Gets the shape bounds in master units.</summary>
        public LegacyPptBounds Bounds { get; }

        /// <summary>Gets the flattened text content.</summary>
        public string Text { get; }

        /// <summary>Gets the placeholder kind.</summary>
        public LegacyPptPlaceholderKind PlaceholderKind { get; }

        /// <summary>Gets decoded OfficeArt fill, line, and shadow properties.</summary>
        public OfficeArtShapeStyle Style { get; }

        /// <summary>Gets the resolved solid fill color as RRGGBB, when available.</summary>
        public string? FillColor { get; }

        /// <summary>Gets the resolved line color as RRGGBB, when available.</summary>
        public string? LineColor { get; }

        /// <summary>Gets the one-based OfficeArt BStore index referenced by the picture frame.</summary>
        public int? PictureStoreIndex { get; }

        /// <summary>Gets the resolved OfficeArt picture entry, when available.</summary>
        public OfficeArtBlipStoreEntry? Picture { get; }

        /// <summary>Gets the coordinate system used by child anchors when this is a group shape.</summary>
        public LegacyPptBounds? GroupCoordinateBounds { get; }

        /// <summary>Gets nested shapes in drawing order when this is a group shape.</summary>
        public IReadOnlyList<LegacyPptShape> Children { get; }
    }
}
