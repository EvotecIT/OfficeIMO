namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Identifies the page-size category stored by a binary PowerPoint document.</summary>
    public enum LegacyPptSlideSizeType : ushort {
        /// <summary>A computer-screen presentation.</summary>
        Screen = 0,
        /// <summary>Letter paper.</summary>
        LetterPaper = 1,
        /// <summary>A4 paper.</summary>
        A4Paper = 2,
        /// <summary>35mm photo slides.</summary>
        Film35Mm = 3,
        /// <summary>Overhead-projector slides.</summary>
        Overhead = 4,
        /// <summary>A banner.</summary>
        Banner = 5,
        /// <summary>A custom size.</summary>
        Custom = 6
    }

    /// <summary>Represents the complete fixed-width DocumentAtom settings.</summary>
    public sealed class LegacyPptDocumentSettings {
        internal LegacyPptDocumentSettings(int slideWidth, int slideHeight, int notesWidth,
            int notesHeight, int serverZoomNumerator, int serverZoomDenominator,
            uint notesMasterPersistId, uint handoutMasterPersistId, ushort firstSlideNumber,
            ushort rawSlideSizeType, bool saveWithFonts, bool omitTitlePlaceholders,
            bool rightToLeft, bool showComments) {
            SlideWidth = slideWidth;
            SlideHeight = slideHeight;
            NotesWidth = notesWidth;
            NotesHeight = notesHeight;
            ServerZoomNumerator = serverZoomNumerator;
            ServerZoomDenominator = serverZoomDenominator;
            NotesMasterPersistId = notesMasterPersistId;
            HandoutMasterPersistId = handoutMasterPersistId;
            FirstSlideNumber = firstSlideNumber;
            RawSlideSizeType = rawSlideSizeType;
            SlideSizeType = Enum.IsDefined(typeof(LegacyPptSlideSizeType), rawSlideSizeType)
                ? (LegacyPptSlideSizeType)rawSlideSizeType
                : null;
            SaveWithFonts = saveWithFonts;
            OmitTitlePlaceholders = omitTitlePlaceholders;
            RightToLeft = rightToLeft;
            ShowComments = showComments;
        }

        /// <summary>Gets the presentation width in master units.</summary>
        public int SlideWidth { get; }
        /// <summary>Gets the presentation height in master units.</summary>
        public int SlideHeight { get; }
        /// <summary>Gets the notes and handout width in master units.</summary>
        public int NotesWidth { get; }
        /// <summary>Gets the notes and handout height in master units.</summary>
        public int NotesHeight { get; }
        /// <summary>Gets the OLE server-zoom numerator.</summary>
        public int ServerZoomNumerator { get; }
        /// <summary>Gets the OLE server-zoom denominator.</summary>
        public int ServerZoomDenominator { get; }
        /// <summary>Gets the notes-master persist reference.</summary>
        public uint NotesMasterPersistId { get; }
        /// <summary>Gets the handout-master persist reference.</summary>
        public uint HandoutMasterPersistId { get; }
        /// <summary>Gets the first displayed slide number.</summary>
        public ushort FirstSlideNumber { get; }
        /// <summary>Gets the raw binary slide-size category.</summary>
        public ushort RawSlideSizeType { get; }
        /// <summary>Gets the typed slide-size category, or null when undefined.</summary>
        public LegacyPptSlideSizeType? SlideSizeType { get; }
        /// <summary>Gets whether the source requests embedded font programs.</summary>
        public bool SaveWithFonts { get; }
        /// <summary>Gets whether special placeholders are omitted from title slides.</summary>
        public bool OmitTitlePlaceholders { get; }
        /// <summary>Gets whether the presentation UI is optimized for right-to-left languages.</summary>
        public bool RightToLeft { get; }
        /// <summary>Gets whether legacy comments are displayed.</summary>
        public bool ShowComments { get; }
    }
}
