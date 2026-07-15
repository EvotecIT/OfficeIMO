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

        /// <summary>A recognized OfficeArt shape that is preserved only as metadata.</summary>
        Unsupported
    }

    /// <summary>Identifies the relevant legacy placeholder types.</summary>
    public enum LegacyPptPlaceholderKind : byte {
        /// <summary>The shape is not a placeholder.</summary>
        None = 0x00,
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
        VerticalBody = 0x12
    }

    /// <summary>Represents a shape decoded from a binary PowerPoint slide.</summary>
    public sealed class LegacyPptShape {
        internal LegacyPptShape(LegacyPptShapeKind kind, ushort officeArtShapeType, uint shapeId,
            long recordOffset, LegacyPptBounds bounds, string text, LegacyPptPlaceholderKind placeholderKind) {
            Kind = kind;
            OfficeArtShapeType = officeArtShapeType;
            ShapeId = shapeId;
            RecordOffset = recordOffset;
            Bounds = bounds;
            Text = text ?? string.Empty;
            PlaceholderKind = placeholderKind;
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
    }
}
