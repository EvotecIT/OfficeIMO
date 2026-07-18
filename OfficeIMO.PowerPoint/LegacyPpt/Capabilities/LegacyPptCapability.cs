namespace OfficeIMO.PowerPoint.LegacyPpt.Capabilities {
    /// <summary>Identifies a PowerPoint capability covered by the binary-format compatibility contract.</summary>
    public enum LegacyPptFeature {
        /// <summary>Normal document lifecycle APIs.</summary>
        FileLifecycle,
        /// <summary>Legacy presentation, template, and slide-show variants.</summary>
        FileVariants,
        /// <summary>Presentation page dimensions.</summary>
        SlideSize,
        /// <summary>Slide containers and content.</summary>
        Slides,
        /// <summary>Slide display order.</summary>
        SlideOrder,
        /// <summary>Hidden-slide state.</summary>
        SlideVisibility,
        /// <summary>Modern presentation sections.</summary>
        Sections,
        /// <summary>Named custom slide shows.</summary>
        CustomShows,
        /// <summary>Presentation master records.</summary>
        Masters,
        /// <summary>Slide layout inheritance.</summary>
        Layouts,
        /// <summary>Theme-like design information.</summary>
        Themes,
        /// <summary>Legacy color schemes.</summary>
        ColorSchemes,
        /// <summary>Placeholder identity and inheritance.</summary>
        Placeholders,
        /// <summary>Master and slide backgrounds.</summary>
        Backgrounds,
        /// <summary>Plain editable text.</summary>
        PlainText,
        /// <summary>Character-level text formatting.</summary>
        RichText,
        /// <summary>Paragraph-level text formatting.</summary>
        ParagraphFormatting,
        /// <summary>Bullets, numbering, and list levels.</summary>
        BulletsAndNumbering,
        /// <summary>Text wrapping and automatic fitting.</summary>
        TextAutoFit,
        /// <summary>Text and shape hyperlinks.</summary>
        Hyperlinks,
        /// <summary>Shape and slide actions.</summary>
        Actions,
        /// <summary>OfficeArt AutoShapes.</summary>
        AutoShapes,
        /// <summary>OfficeArt connectors.</summary>
        Connectors,
        /// <summary>Nested shape groups.</summary>
        Groups,
        /// <summary>Shape position, size, rotation, and flip.</summary>
        ShapeTransforms,
        /// <summary>Shape fills and outlines.</summary>
        ShapeStyles,
        /// <summary>Legacy shape effects.</summary>
        ShapeEffects,
        /// <summary>Raster picture formats.</summary>
        RasterPictures,
        /// <summary>Windows and enhanced metafile pictures.</summary>
        MetafilePictures,
        /// <summary>Picture crop and transform settings.</summary>
        PictureCrop,
        /// <summary>Native presentation tables.</summary>
        Tables,
        /// <summary>Legacy embedded charts.</summary>
        Charts,
        /// <summary>Modern SmartArt diagrams.</summary>
        SmartArt,
        /// <summary>Editable speaker-note text.</summary>
        SpeakerNotes,
        /// <summary>Rich notes-page content.</summary>
        RichNotes,
        /// <summary>Header, footer, date, and slide-number settings.</summary>
        HeadersAndFooters,
        /// <summary>Presentation comments.</summary>
        Comments,
        /// <summary>Modern threaded comments, replies, status, and anchors.</summary>
        ModernComments,
        /// <summary>Slide transition settings.</summary>
        Transitions,
        /// <summary>Embedded sounds used by transitions and interactive actions.</summary>
        TransitionAndActionSounds,
        /// <summary>Animation and timing trees.</summary>
        Animations,
        /// <summary>Embedded WAV audio content.</summary>
        Media,
        /// <summary>Embedded OLE objects.</summary>
        EmbeddedOle,
        /// <summary>ActiveX controls.</summary>
        ActiveX,
        /// <summary>VBA project storage.</summary>
        VbaProjects,
        /// <summary>Built-in document properties.</summary>
        BuiltInProperties,
        /// <summary>Custom document properties.</summary>
        CustomProperties,
        /// <summary>Open XML custom XML parts.</summary>
        CustomXml,
        /// <summary>Legacy binary encryption.</summary>
        Encryption,
        /// <summary>Legacy digital signatures.</summary>
        DigitalSignatures,
        /// <summary>Alternative text and object naming.</summary>
        AccessibilityMetadata,
        /// <summary>Opaque or unknown records, streams, and storages.</summary>
        UnknownRecordsAndStreams,
        /// <summary>Linked OLE objects and their cached storages.</summary>
        LinkedOle,
        /// <summary>Linked and device-based legacy audio and video.</summary>
        LinkedMedia,
        /// <summary>Modern embedded video content with no binary equivalent.</summary>
        EmbeddedVideo
    }

    /// <summary>Describes how a feature can be represented in PowerPoint 97-2003 binary files.</summary>
    public enum LegacyPptRepresentability {
        /// <summary>The binary format has a native representation that can be made editable.</summary>
        Native,

        /// <summary>The feature can be represented through a documented approximation or flattening.</summary>
        Approximation,

        /// <summary>The content can be retained as opaque binary data but is not safely editable.</summary>
        Opaque,

        /// <summary>The feature has no supported representation in the legacy format.</summary>
        NotRepresentable
    }

    /// <summary>Describes the current support state for one binary PowerPoint capability lane.</summary>
    public enum LegacyPptCapabilityState {
        /// <summary>The capability is decoded or encoded natively and remains editable.</summary>
        Native,

        /// <summary>The capability is retained without being projected to an editable model.</summary>
        Preserved,

        /// <summary>The capability is deliberately converted to a documented legacy approximation.</summary>
        Converted,

        /// <summary>The operation is deliberately refused to prevent unsupported loss.</summary>
        Blocked,

        /// <summary>The capability is part of the parity target but is not implemented yet.</summary>
        Planned
    }

    /// <summary>Identifies one direction of the binary PowerPoint compatibility contract.</summary>
    public enum LegacyPptCapabilityLane {
        /// <summary>Binary input projected into the normal editable presentation model.</summary>
        ImportToEditableModel,
        /// <summary>A new binary presentation authored from the normal presentation model.</summary>
        NewBinaryWrite,
        /// <summary>An imported binary presentation edited and saved back as binary.</summary>
        BinaryRoundTrip,
        /// <summary>An Open XML presentation converted to a legacy binary presentation.</summary>
        PptxToBinary
    }

    /// <summary>One machine-readable row in the binary PowerPoint parity contract.</summary>
    public sealed class LegacyPptCapability {
        internal LegacyPptCapability(LegacyPptFeature feature, string category, string description,
            LegacyPptRepresentability representability, LegacyPptCapabilityState importToEditableModel,
            LegacyPptCapabilityState newBinaryWrite, LegacyPptCapabilityState binaryRoundTrip,
            LegacyPptCapabilityState pptxToBinary, string note) {
            Feature = feature;
            Category = category ?? throw new ArgumentNullException(nameof(category));
            Description = description ?? throw new ArgumentNullException(nameof(description));
            Representability = representability;
            ImportToEditableModel = importToEditableModel;
            NewBinaryWrite = newBinaryWrite;
            BinaryRoundTrip = binaryRoundTrip;
            PptxToBinary = pptxToBinary;
            Note = note ?? string.Empty;
        }

        /// <summary>Gets the stable feature identifier.</summary>
        public LegacyPptFeature Feature { get; }

        /// <summary>Gets the feature category used by reports.</summary>
        public string Category { get; }

        /// <summary>Gets the user-facing feature description.</summary>
        public string Description { get; }

        /// <summary>Gets how the legacy format can represent the feature.</summary>
        public LegacyPptRepresentability Representability { get; }

        /// <summary>Gets the current binary-to-editable-model import state.</summary>
        public LegacyPptCapabilityState ImportToEditableModel { get; }

        /// <summary>Gets the current new-binary-document authoring state.</summary>
        public LegacyPptCapabilityState NewBinaryWrite { get; }

        /// <summary>Gets the current imported-binary-to-binary round-trip state.</summary>
        public LegacyPptCapabilityState BinaryRoundTrip { get; }

        /// <summary>Gets the current PPTX-to-binary conversion state.</summary>
        public LegacyPptCapabilityState PptxToBinary { get; }

        /// <summary>Gets an important compatibility note or limitation.</summary>
        public string Note { get; }

        /// <summary>Gets whether any compatibility lane still requires implementation work.</summary>
        public bool HasRemainingParityWork => ImportToEditableModel == LegacyPptCapabilityState.Planned
            || NewBinaryWrite == LegacyPptCapabilityState.Planned
            || BinaryRoundTrip == LegacyPptCapabilityState.Planned
            || PptxToBinary == LegacyPptCapabilityState.Planned;

        /// <summary>Gets the current state for a requested compatibility lane.</summary>
        public LegacyPptCapabilityState GetState(LegacyPptCapabilityLane lane) {
            switch (lane) {
                case LegacyPptCapabilityLane.ImportToEditableModel: return ImportToEditableModel;
                case LegacyPptCapabilityLane.NewBinaryWrite: return NewBinaryWrite;
                case LegacyPptCapabilityLane.BinaryRoundTrip: return BinaryRoundTrip;
                case LegacyPptCapabilityLane.PptxToBinary: return PptxToBinary;
                default: throw new ArgumentOutOfRangeException(nameof(lane));
            }
        }
    }
}
