namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Specifies the layout hint stored in a binary PowerPoint SlideAtom.</summary>
    public enum LegacyPptSlideLayoutType : uint {
        /// <summary>One centered title and one subtitle placeholder.</summary>
        TitleSlide = 0x00000000,
        /// <summary>One title and one body or content placeholder.</summary>
        TitleBody = 0x00000001,
        /// <summary>A title-master layout with one centered title and one subtitle placeholder.</summary>
        MasterTitle = 0x00000002,
        /// <summary>One title placeholder.</summary>
        TitleOnly = 0x00000007,
        /// <summary>One title and two horizontally arranged content placeholders.</summary>
        TwoColumns = 0x00000008,
        /// <summary>One title and two vertically arranged content placeholders.</summary>
        TwoRows = 0x00000009,
        /// <summary>One title with one left and two right content placeholders.</summary>
        ColumnTwoRows = 0x0000000A,
        /// <summary>One title with two left and one right content placeholders.</summary>
        TwoRowsColumn = 0x0000000B,
        /// <summary>One title with two upper and one lower content placeholders.</summary>
        TwoColumnsRow = 0x0000000D,
        /// <summary>One title and four content placeholders.</summary>
        FourObjects = 0x0000000E,
        /// <summary>One large content placeholder.</summary>
        BigObject = 0x0000000F,
        /// <summary>A blank or custom layout.</summary>
        Blank = 0x00000010,
        /// <summary>One vertical title and one vertical body placeholder.</summary>
        VerticalTitleBody = 0x00000011,
        /// <summary>One vertical title and two vertically oriented content placeholders.</summary>
        VerticalTwoRows = 0x00000012
    }

    /// <summary>Specifies a placeholder's preferred size relative to the master body placeholder.</summary>
    public enum LegacyPptPlaceholderSize : byte {
        /// <summary>Full master-body size.</summary>
        Full = 0x00,
        /// <summary>Half master-body size.</summary>
        Half = 0x01,
        /// <summary>Quarter master-body size.</summary>
        Quarter = 0x02
    }

    /// <summary>Represents the identity, kind, and preferred size from a PlaceholderAtom.</summary>
    public sealed class LegacyPptPlaceholder {
        internal LegacyPptPlaceholder(int position, LegacyPptPlaceholderKind kind,
            LegacyPptPlaceholderSize size) {
            Position = position;
            Kind = kind;
            Size = size;
        }

        /// <summary>Gets the placeholder identifier, unique within its slide when the source is valid.</summary>
        public int Position { get; }

        /// <summary>Gets the placeholder kind.</summary>
        public LegacyPptPlaceholderKind Kind { get; }

        /// <summary>Gets the preferred placeholder size.</summary>
        public LegacyPptPlaceholderSize Size { get; }
    }
}
