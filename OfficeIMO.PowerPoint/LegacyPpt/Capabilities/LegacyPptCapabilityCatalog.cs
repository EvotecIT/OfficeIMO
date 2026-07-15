using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint.LegacyPpt.Capabilities {
    /// <summary>
    /// Provides the versioned source of truth for binary PowerPoint import, authoring, round-trip, and conversion support.
    /// </summary>
    public static partial class LegacyPptCapabilityCatalog {
        private static readonly IReadOnlyList<LegacyPptCapability> CapabilityRows =
            new ReadOnlyCollection<LegacyPptCapability>(new[] {
                Native(LegacyPptFeature.FileLifecycle, "Lifecycle", "Load, save, save-copy, streams, bytes, and async flows."),
                Native(LegacyPptFeature.FileVariants, "Lifecycle", ".ppt, .pot, and .pps routing and extension semantics."),
                Native(LegacyPptFeature.SlideSize, "Structure", "Presentation page dimensions."),
                Native(LegacyPptFeature.Slides, "Structure", "Slide containers and editable slide content."),
                Native(LegacyPptFeature.SlideOrder, "Structure", "Display-order slide directory entries."),
                Planned(LegacyPptFeature.SlideVisibility, "Structure", "Hidden-slide state."),
                Blocked(LegacyPptFeature.Sections, "Structure", "Modern presentation sections.",
                    "PowerPoint 97-2003 has no native section model."),
                Planned(LegacyPptFeature.CustomShows, "Structure", "Named custom slide shows."),
                Planned(LegacyPptFeature.Masters, "Design", "Main, title, notes, and handout masters."),
                Planned(LegacyPptFeature.Layouts, "Design", "Slide layout and master inheritance."),
                Planned(LegacyPptFeature.Themes, "Design", "Theme fonts, fills, lines, and effects.",
                    LegacyPptRepresentability.Approximation),
                Planned(LegacyPptFeature.ColorSchemes, "Design", "Legacy master and slide color schemes."),
                Planned(LegacyPptFeature.Placeholders, "Design", "Placeholder identity, type, size, and inheritance.",
                    "The bootstrap currently projects common text placeholders only."),
                Planned(LegacyPptFeature.Backgrounds, "Design", "Master and slide backgrounds."),
                Native(LegacyPptFeature.PlainText, "Text", "Unicode and byte text projected as editable plain text."),
                Planned(LegacyPptFeature.RichText, "Text", "Character runs, fonts, sizes, colors, and emphasis."),
                Planned(LegacyPptFeature.ParagraphFormatting, "Text", "Alignment, indentation, spacing, tabs, and margins."),
                Planned(LegacyPptFeature.BulletsAndNumbering, "Text", "Bullet glyphs, pictures, numbering, and levels."),
                Planned(LegacyPptFeature.TextAutoFit, "Text", "Text fitting, wrapping, and text-box inset behavior."),
                Planned(LegacyPptFeature.Hyperlinks, "Interaction", "Text and shape hyperlinks."),
                Planned(LegacyPptFeature.Actions, "Interaction", "Shape actions, slide jumps, and programs."),
                Planned(LegacyPptFeature.AutoShapes, "Drawing", "OfficeArt AutoShape geometry.",
                    "The bootstrap writer currently supports rectangle, ellipse, and line only."),
                Planned(LegacyPptFeature.Connectors, "Drawing", "Connector shapes and connection sites."),
                Planned(LegacyPptFeature.Groups, "Drawing", "Nested drawing groups and child coordinate systems."),
                Planned(LegacyPptFeature.ShapeTransforms, "Drawing", "Position, size, rotation, flip, and z-order.",
                    "Imported mapped shapes preserve position and size edits through incremental binary records; rotation, flip, and z-order remain planned."),
                Planned(LegacyPptFeature.ShapeStyles, "Drawing", "Fill, outline, transparency, and shape properties."),
                Planned(LegacyPptFeature.ShapeEffects, "Drawing", "Shadows and legacy OfficeArt effects."),
                Planned(LegacyPptFeature.RasterPictures, "Images", "PNG, JPEG, DIB, and bitmap BLIP records."),
                Planned(LegacyPptFeature.MetafilePictures, "Images", "WMF and EMF BLIP records."),
                Planned(LegacyPptFeature.PictureCrop, "Images", "Picture crop, transform, and recolor properties."),
                Planned(LegacyPptFeature.Tables, "Content", "Native OfficeArt tables and cell formatting."),
                Planned(LegacyPptFeature.Charts, "Content", "Legacy Microsoft Graph and embedded chart objects."),
                Planned(LegacyPptFeature.SmartArt, "Content", "SmartArt diagrams.", LegacyPptRepresentability.Approximation,
                    "SmartArt requires an explicit conversion to grouped OfficeArt or a static visual."),
                Planned(LegacyPptFeature.SpeakerNotes, "Presentation", "Editable speaker notes.",
                    "The bootstrap reader imports plain note text but does not write notes."),
                Planned(LegacyPptFeature.RichNotes, "Presentation", "Notes-page drawings and formatting."),
                Planned(LegacyPptFeature.HeadersAndFooters, "Presentation", "Date, footer, header, and slide-number settings."),
                Planned(LegacyPptFeature.Comments, "Review", "Legacy comment authors and comment records."),
                Planned(LegacyPptFeature.Transitions, "Presentation", "Slide transitions, speed, and advance settings."),
                Planned(LegacyPptFeature.Animations, "Presentation", "Animation and timing trees."),
                Planned(LegacyPptFeature.Media, "Media", "Embedded and linked audio/video with playback settings."),
                Planned(LegacyPptFeature.EmbeddedOle, "Embedded", "Embedded and linked OLE objects.",
                    LegacyPptRepresentability.Opaque),
                Planned(LegacyPptFeature.ActiveX, "Embedded", "ActiveX controls and associated storages.",
                    LegacyPptRepresentability.Opaque),
                Planned(LegacyPptFeature.VbaProjects, "Embedded", "VBA project storage.",
                    LegacyPptRepresentability.Opaque),
                Planned(LegacyPptFeature.BuiltInProperties, "Metadata", "Summary and document-summary properties."),
                Planned(LegacyPptFeature.CustomProperties, "Metadata", "Custom document properties."),
                Blocked(LegacyPptFeature.CustomXml, "Metadata", "Open XML custom XML parts.",
                    "PowerPoint 97-2003 has no equivalent package-part representation."),
                Planned(LegacyPptFeature.Encryption, "Security", "Binary RC4 CryptoAPI password encryption."),
                Planned(LegacyPptFeature.DigitalSignatures, "Security", "Legacy digital-signature storages.",
                    LegacyPptRepresentability.Opaque,
                    "Unmodified signatures can be preserved; edits invalidate signature integrity."),
                Planned(LegacyPptFeature.AccessibilityMetadata, "Accessibility", "Alternative text and object names."),
                new LegacyPptCapability(LegacyPptFeature.UnknownRecordsAndStreams, "Preservation",
                    "Unknown live records and compound streams.", LegacyPptRepresentability.Opaque,
                    LegacyPptCapabilityState.Preserved, LegacyPptCapabilityState.Blocked,
                    LegacyPptCapabilityState.Planned, LegacyPptCapabilityState.Blocked,
                    "No-op saves retain the exact package. Geometry-only edits append a UserEdit and preserve untouched records and streams; broader edited preservation remains planned.")
            });

        private static readonly IReadOnlyDictionary<LegacyPptFeature, LegacyPptCapability> CapabilityByFeature =
            new ReadOnlyDictionary<LegacyPptFeature, LegacyPptCapability>(CapabilityRows.ToDictionary(row => row.Feature));

        /// <summary>Gets the contract schema version.</summary>
        public static int SchemaVersion => 1;

        /// <summary>Gets every capability row in stable feature order.</summary>
        public static IReadOnlyList<LegacyPptCapability> Capabilities => CapabilityRows;

        /// <summary>Gets whether any capability lane still contains planned parity work.</summary>
        public static bool HasRemainingParityWork => CapabilityRows.Any(row => row.HasRemainingParityWork);

        /// <summary>Gets all capability rows that still require implementation.</summary>
        public static IReadOnlyList<LegacyPptCapability> RemainingParityWork => CapabilityRows
            .Where(row => row.HasRemainingParityWork)
            .ToArray();

        /// <summary>Gets the contract row for a feature.</summary>
        public static LegacyPptCapability Get(LegacyPptFeature feature) {
            if (!CapabilityByFeature.TryGetValue(feature, out LegacyPptCapability? capability)) {
                throw new ArgumentOutOfRangeException(nameof(feature), feature, "The feature is not present in the capability contract.");
            }
            return capability;
        }

        private static LegacyPptCapability Native(LegacyPptFeature feature, string category, string description) =>
            new LegacyPptCapability(feature, category, description, LegacyPptRepresentability.Native,
                LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Native,
                LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Native, string.Empty);

        private static LegacyPptCapability Planned(LegacyPptFeature feature, string category, string description,
            string note = "") => Planned(feature, category, description, LegacyPptRepresentability.Native, note);

        private static LegacyPptCapability Planned(LegacyPptFeature feature, string category, string description,
            LegacyPptRepresentability representability, string note = "") =>
            new LegacyPptCapability(feature, category, description, representability,
                LegacyPptCapabilityState.Planned, LegacyPptCapabilityState.Planned,
                LegacyPptCapabilityState.Planned, LegacyPptCapabilityState.Planned, note);

        private static LegacyPptCapability Blocked(LegacyPptFeature feature, string category, string description,
            string note) =>
            new LegacyPptCapability(feature, category, description, LegacyPptRepresentability.NotRepresentable,
                LegacyPptCapabilityState.Blocked, LegacyPptCapabilityState.Blocked,
                LegacyPptCapabilityState.Blocked, LegacyPptCapabilityState.Blocked, note);
    }
}
