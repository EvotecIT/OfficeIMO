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

        /// <summary>A native OfficeArt table group with semantic rows and cells.</summary>
        Table,

        /// <summary>An OfficeArt picture frame with importable image data.</summary>
        Picture,

        /// <summary>An embedded OLE compound object.</summary>
        OleObject,

        /// <summary>Embedded WAV audio projected as an editable media shape.</summary>
        Media,

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
            long recordOffset, LegacyPptBounds bounds, string text, LegacyPptPlaceholder? placeholder,
            OfficeArtShapeStyle style, string? fillColor, string? lineColor,
            int? pictureStoreIndex = null, OfficeArtBlipStoreEntry? picture = null,
            OfficeArtShapeTransform? transform = null,
            LegacyPptBounds? groupCoordinateBounds = null,
            IReadOnlyList<LegacyPptShape>? children = null,
            byte? tableStyleFlags = null,
            string? shadowColor = null,
            LegacyPptTextBody? textBody = null,
            IReadOnlyList<LegacyPptInteraction>? interactions = null,
            LegacyPptAnimation? animation = null,
            LegacyPptEmbeddedOleObject? oleObject = null,
            LegacyPptLinkedOleObject? linkedOleObject = null,
            LegacyPptActiveXControl? activeXControl = null,
            LegacyPptMedia? media = null,
            string? pictureTransparentColor = null,
            string? pictureRecolorColor = null,
            string? fillBackColor = null,
            IReadOnlyList<LegacyPptGradientStop>? fillGradientStops = null) {
            Kind = kind;
            OfficeArtShapeType = officeArtShapeType;
            ShapeId = shapeId;
            RecordOffset = recordOffset;
            Bounds = bounds;
            Text = text ?? string.Empty;
            TextBody = textBody ?? LegacyPptTextBody.Plain(Text);
            Placeholder = placeholder;
            Style = style ?? throw new ArgumentNullException(nameof(style));
            FillColor = fillColor;
            FillBackColor = fillBackColor;
            FillGradientStops = new ReadOnlyCollection<LegacyPptGradientStop>(
                fillGradientStops?.ToArray() ?? Array.Empty<LegacyPptGradientStop>());
            LineColor = lineColor;
            PictureStoreIndex = pictureStoreIndex;
            Picture = picture;
            PictureTransparentColor = pictureTransparentColor;
            PictureRecolorColor = pictureRecolorColor;
            Transform = transform ?? OfficeArtShapeTransform.Decode(0);
            Geometry = OfficeArtShapeGeometry.Decode(style.Properties);
            PictureProperties = OfficeArtPictureProperties.Decode(style.Properties);
            Metadata = OfficeArtShapeMetadata.Decode(style.Properties);
            TextFrame = LegacyPptTextFrameProperties.Decode(style.Properties);
            ShadowColor = shadowColor;
            GroupCoordinateBounds = groupCoordinateBounds;
            Children = new ReadOnlyCollection<LegacyPptShape>(
                children?.ToArray() ?? Array.Empty<LegacyPptShape>());
            Table = kind == LegacyPptShapeKind.Group
                ? LegacyPptTable.TryCreate(Style, Children, tableStyleFlags)
                : null;
            if (Table != null) Kind = LegacyPptShapeKind.Table;
            Interactions = new ReadOnlyCollection<LegacyPptInteraction>(
                interactions?.ToArray() ?? Array.Empty<LegacyPptInteraction>());
            Animation = animation;
            OleObject = oleObject;
            LinkedOleObject = linkedOleObject;
            ActiveXControl = activeXControl;
            Media = media;
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

        /// <summary>Gets decoded character-run and paragraph-style information for the shape text.</summary>
        public LegacyPptTextBody TextBody { get; }

        /// <summary>Gets the decoded placeholder identity and size, when this is a placeholder shape.</summary>
        public LegacyPptPlaceholder? Placeholder { get; }

        /// <summary>Gets the placeholder kind.</summary>
        public LegacyPptPlaceholderKind PlaceholderKind => Placeholder?.Kind
            ?? LegacyPptPlaceholderKind.None;

        /// <summary>Gets decoded OfficeArt fill, line, and shadow properties.</summary>
        public OfficeArtShapeStyle Style { get; }

        /// <summary>Gets the resolved foreground or first gradient fill color as RRGGBB.</summary>
        public string? FillColor { get; }

        /// <summary>Gets the resolved background or final gradient fill color as RRGGBB.</summary>
        public string? FillBackColor { get; }

        /// <summary>Gets resolved custom OfficeArt gradient stops.</summary>
        public IReadOnlyList<LegacyPptGradientStop> FillGradientStops { get; }

        /// <summary>Gets the resolved line color as RRGGBB, when available.</summary>
        public string? LineColor { get; }

        /// <summary>Gets the resolved shadow color as RRGGBB, when available.</summary>
        public string? ShadowColor { get; }

        /// <summary>Gets the one-based OfficeArt BStore index referenced by the picture frame.</summary>
        public int? PictureStoreIndex { get; }

        /// <summary>Gets the resolved OfficeArt picture entry, when available.</summary>
        public OfficeArtBlipStoreEntry? Picture { get; }

        /// <summary>Gets the resolved picture transparent color as RRGGBB.</summary>
        public string? PictureTransparentColor { get; }

        /// <summary>Gets the resolved whole-picture recolor color as RRGGBB.</summary>
        public string? PictureRecolorColor { get; }

        /// <summary>Gets decoded rotation and mirroring state.</summary>
        public OfficeArtShapeTransform Transform { get; }

        /// <summary>Gets raw shape-specific OfficeArt geometry adjustment values.</summary>
        public OfficeArtShapeGeometry Geometry { get; }

        /// <summary>Gets decoded OfficeArt picture-frame crop properties.</summary>
        public OfficeArtPictureProperties PictureProperties { get; }

        /// <summary>Gets decoded object name and description metadata.</summary>
        public OfficeArtShapeMetadata Metadata { get; }

        /// <summary>Gets decoded OfficeArt text-frame properties.</summary>
        public LegacyPptTextFrameProperties TextFrame { get; }

        /// <summary>Gets the coordinate system used by child anchors when this is a group shape.</summary>
        public LegacyPptBounds? GroupCoordinateBounds { get; }

        /// <summary>Gets nested shapes in drawing order when this is a group shape.</summary>
        public IReadOnlyList<LegacyPptShape> Children { get; }

        /// <summary>Gets the decoded native table when this group carries OfficeArt table semantics.</summary>
        public LegacyPptTable? Table { get; }

        /// <summary>Gets shape-level click and mouse-over interactions.</summary>
        public IReadOnlyList<LegacyPptInteraction> Interactions { get; }

        /// <summary>Gets the classic shape or text animation, when present and valid.</summary>
        public LegacyPptAnimation? Animation { get; }

        /// <summary>Gets the embedded OLE object referenced by this shape.</summary>
        public LegacyPptEmbeddedOleObject? OleObject { get; }

        /// <summary>Gets the preserve-only linked OLE object referenced by this shape.</summary>
        public LegacyPptLinkedOleObject? LinkedOleObject { get; }

        /// <summary>Gets the preserve-only ActiveX control referenced by this shape.</summary>
        public LegacyPptActiveXControl? ActiveXControl { get; }

        /// <summary>Gets the binary audio or movie object referenced by this shape.</summary>
        public LegacyPptMedia? Media { get; }

        /// <summary>Gets the shape-level click interaction, when present.</summary>
        public LegacyPptInteraction? ClickInteraction => Interactions.FirstOrDefault(
            interaction => interaction.Trigger == LegacyPptInteractionTrigger.MouseClick);

        /// <summary>Gets the shape-level mouse-over interaction, when present.</summary>
        public LegacyPptInteraction? MouseOverInteraction => Interactions.FirstOrDefault(
            interaction => interaction.Trigger == LegacyPptInteractionTrigger.MouseOver);
    }
}
