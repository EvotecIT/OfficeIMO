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
                Native(LegacyPptFeature.SlideVisibility, "Structure", "Hidden-slide state."),
                Blocked(LegacyPptFeature.Sections, "Structure", "Modern presentation sections.",
                    "PowerPoint 97-2003 has no native section model."),
                new LegacyPptCapability(LegacyPptFeature.CustomShows, "Structure",
                    "Named custom slide shows.", LegacyPptRepresentability.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Native,
                    "NamedShows records decode to typed ordered slide-id sequences and project to native custom-show lists. Fresh PPT/POT/PPS writing and imported add/edit/remove preserve custom-show membership and identity incrementally. Custom-show actions, including show-and-return, project and round-trip natively."),
                Planned(LegacyPptFeature.Masters, "Design", "Main, title, notes, and handout masters.",
                    "Main and title master records, supported shapes, and base title/body/other text styles through five levels are projected into native master/layout parts. DocumentAtom notes and handout master references, schemes, shapes, placeholders, and connector rules project into native notesMaster and handoutMaster parts. Later-version text-style extensions and binary master editing remain planned."),
                new LegacyPptCapability(LegacyPptFeature.Layouts, "Design",
                    "Slide layout and master inheritance.", LegacyPptRepresentability.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Planned,
                    LegacyPptCapabilityState.Preserved, LegacyPptCapabilityState.Planned,
                    "All 14 SlideLayoutType values and their eight-slot placeholder signatures decode to typed metadata and distinct native Open XML layouts under the referenced master. Supported PPTX layouts write native binary hints for fresh and incrementally appended slides, and imported relationship or layout edits are loss-blocked. Regenerating arbitrary custom binary layout definitions, inherited geometry, and layout styling remains planned."),
                Planned(LegacyPptFeature.Themes, "Design", "Theme fonts, fills, lines, and effects.",
                    LegacyPptRepresentability.Approximation),
                Planned(LegacyPptFeature.ColorSchemes, "Design", "Legacy master and slide color schemes.",
                    LegacyPptRepresentability.Approximation,
                    "All eight legacy scheme colors are decoded and mapped to native master or slide theme slots; exact legacy-to-Open-XML slot semantics and edited binary writing remain planned."),
                new LegacyPptCapability(LegacyPptFeature.Placeholders, "Design",
                    "Placeholder identity, type, size, and inheritance.",
                    LegacyPptRepresentability.Native, LegacyPptCapabilityState.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Preserved,
                    LegacyPptCapabilityState.Native,
                    "PlaceholderAtom position, all legacy kinds, full/half/quarter size, and vertical orientation project to native placeholder index/type/size/orientation on masters, generated layouts, and slides. Fresh and incrementally appended binary slides encode the same fields; edits to imported placeholder contracts are loss-blocked. Object payloads are tracked by their separate capability rows."),
                Planned(LegacyPptFeature.Backgrounds, "Design", "Master and slide backgrounds.",
                    "OfficeArt background shapes are distinguished from normal drawing content. Solid, no-fill, inherited, multi-stop linear/path gradient, texture, and picture backgrounds project to native Open XML fills on slides and all master types; gradient focus, exact pattern semantics, and edited binary writing remain planned."),
                new LegacyPptCapability(LegacyPptFeature.PlainText, "Text",
                    "Unicode and byte text projected as editable plain text.", LegacyPptRepresentability.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Native,
                    LegacyPptCapabilityState.Planned, LegacyPptCapabilityState.Native,
                    "Arbitrary-length edits round-trip for structurally plain text; broader imported text structures remain planned."),
                new LegacyPptCapability(LegacyPptFeature.RichText, "Text",
                    "Character runs, fonts, sizes, colors, and emphasis.", LegacyPptRepresentability.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Planned,
                    LegacyPptCapabilityState.Preserved, LegacyPptCapabilityState.Planned,
                    "StyleTextPropAtom character runs and base TextMasterStyleAtom defaults project bold, italic, underline, point size, direct or resolved scheme color, baseline position, and document font-collection typefaces natively. Embedded font programs, later-version style extensions, and legacy-only character effects remain preserve-only. Unmodified formatting and unrelated supported edits round-trip; formatting edits and fresh binary rich-text authoring remain planned."),
                Planned(LegacyPptFeature.ParagraphFormatting, "Text", "Alignment, indentation, spacing, tabs, and margins.",
                    "StyleTextPropAtom runs, TextRulerAtom records, and five-level base master styles project alignment, margins, first-line indentation, default and explicit tabs, line and paragraph spacing, level, font alignment, RTL direction, East Asian breaking, Latin word wrapping, and hanging punctuation natively. Later-version style extensions, edited binary writing, and fresh authoring remain planned."),
                Planned(LegacyPptFeature.BulletsAndNumbering, "Text", "Bullet glyphs, pictures, numbering, and levels.",
                    "Base character bullets project native enabled/disabled state, glyph, resolved font and color, relative or point size, level, ruler indentation, and five-level master inheritance. Picture bullets, auto-numbering extensions, edited binary writing, and fresh authoring remain planned."),
                Planned(LegacyPptFeature.TextAutoFit, "Text", "Text fitting, wrapping, and text-box inset behavior."),
                Planned(LegacyPptFeature.Hyperlinks, "Interaction", "Text and shape hyperlinks.",
                    "ExObjList, ExHyperlink, and PP9 ExHyperlink9 records decode to typed external and internal-slide targets plus screen tips. Shape and text click/hover interactions project natively, write to new PPT/POT/PPS files, and support preservation-aware add/edit/remove, retargeting, slide reordering, and appended slides. Target frames, custom-show flags, and richer hyperlink metadata remain planned."),
                Planned(LegacyPptFeature.Actions, "Interaction", "Shape actions, slide jumps, and programs.",
                    "InteractiveInfoAtom click and mouse-over records decode to typed action metadata. Built-in slide-show jumps, macro names, program targets, named custom shows with show-and-return, animated-highlight state, and stop-sound state project and write natively, including preservation-aware add/edit/remove. OLE, media, referenced sound, hyperlink-form custom-show metadata, and visited-state actions remain preserve-only or planned."),
                Planned(LegacyPptFeature.AutoShapes, "Drawing", "OfficeArt AutoShape geometry.",
                    "OfficeArt preset geometry with a DrawingML equivalent is projected as an editable native shape, including arrows, callouts, flowcharts, ribbons, stars, and action buttons. All eight signed adjustment slots are decoded without losing their shape-specific meaning; exact round-rectangle and donut adjustments are projected natively. Legacy-only geometries use explicit approximations. Remaining preset adjustments, custom geometry, text warps, and fresh binary authoring remain planned."),
                Planned(LegacyPptFeature.Connectors, "Drawing", "Connector shapes and connection sites.",
                    "Straight, bent, and curved OfficeArt connectors are projected as native editable connection shapes. Solver rules preserve native start/end shape attachments and connection-site indexes, and imported position and size edits round-trip. Editing attachment rules and fresh binary authoring remain planned."),
                Planned(LegacyPptFeature.Groups, "Drawing", "Nested drawing groups and child coordinate systems.",
                    "Nested OfficeArt group hierarchies and child coordinate systems are projected as native editable Open XML groups. Imported outer group geometry edits round-trip; child edits, reparenting, and fresh binary group authoring remain planned."),
                Planned(LegacyPptFeature.ShapeTransforms, "Drawing", "Position, size, rotation, flip, and z-order.",
                    "Position, size, clockwise rotation, and horizontal or vertical mirroring are projected for mapped shapes, pictures, connectors, and nested groups. Position and size edits round-trip through incremental binary records; rotation, flip, child group transforms, and z-order edits remain planned."),
                new LegacyPptCapability(LegacyPptFeature.ShapeStyles, "Drawing",
                    "Fill, outline, transparency, and shape properties.", LegacyPptRepresentability.Native,
                    LegacyPptCapabilityState.Planned, LegacyPptCapabilityState.Planned,
                    LegacyPptCapabilityState.Planned, LegacyPptCapabilityState.Planned,
                    "Shared OfficeArt FOPT decoding projects explicit solid RGB and scheme fills, associated opacity, visibility, line width, preset dashes, joins, caps, arrowheads, and offset-shadow style. Drawing-group defaults, non-solid fills, edited binary style writing, and fresh binary style authoring remain planned."),
                Planned(LegacyPptFeature.ShapeEffects, "Drawing", "Shadows and legacy OfficeArt effects.",
                    "Enabled offset shadows project as native DrawingML outer shadows with resolved color, opacity, signed direction, distance, and available softness. Double, rich, shape, drawing-plane, emboss/engrave, group-level effects, effect editing, and fresh binary authoring remain planned."),
                new LegacyPptCapability(LegacyPptFeature.RasterPictures, "Images",
                    "PNG, JPEG, DIB, and TIFF BLIP records.", LegacyPptRepresentability.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Planned,
                    LegacyPptCapabilityState.Preserved, LegacyPptCapabilityState.Planned,
                    "Shared OfficeArt decoding resolves embedded and delayed picture stores, normalizes DIB data to BMP, and projects referenced slide, master, and layout pictures as editable Open XML image parts. Imported picture position and size edits round-trip; image replacement and fresh binary picture writing remain planned."),
                Planned(LegacyPptFeature.MetafilePictures, "Images", "WMF and EMF BLIP records.",
                    "Shared OfficeArt decoding can extract compressed and uncompressed WMF/EMF payloads for import; corpus interoperability and binary authoring remain planned."),
                Planned(LegacyPptFeature.PictureCrop, "Images", "Picture crop, transform, and recolor properties.",
                    "All four signed OfficeArt crop edges are decoded as 16.16 image fractions and projected natively for slide, master, layout, and grouped pictures, including negative crop-out values. Brightness, contrast, grayscale, and bi-level display state are also projected as native DrawingML effects. Effect thresholds, transparent-color and recolor projection, effect editing, and fresh binary authoring remain planned."),
                Planned(LegacyPptFeature.Tables, "Content", "Native OfficeArt tables and cell formatting."),
                Planned(LegacyPptFeature.Charts, "Content", "Legacy Microsoft Graph and embedded chart objects."),
                Planned(LegacyPptFeature.SmartArt, "Content", "SmartArt diagrams.", LegacyPptRepresentability.Approximation,
                    "SmartArt requires an explicit conversion to grouped OfficeArt or a static visual."),
                new LegacyPptCapability(LegacyPptFeature.SpeakerNotes, "Presentation",
                    "Editable speaker notes.", LegacyPptRepresentability.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Native,
                    "SlideAtom notes identifiers resolve through NotesPersistAtom entries to typed NotesContainer pages. Plain note text projects editably, fresh PPT/POT/PPS writing emits native notes directories and persist objects, and structurally plain imported note edits append preservation-aware records. Length-changing edits to styled note bodies remain loss-blocked."),
                Planned(LegacyPptFeature.RichNotes, "Presentation", "Notes-page drawings and formatting.",
                    "NotesAtom inheritance, schemes, backgrounds, placeholders, rich text, supported drawings, pictures, groups, and connector rules project into native notesSlide parts. Fresh binary rich-notes authoring and arbitrary notes-page editing remain planned."),
                new LegacyPptCapability(LegacyPptFeature.HeadersAndFooters, "Presentation",
                    "Date, footer, header, and slide-number settings.",
                    LegacyPptRepresentability.Native, LegacyPptCapabilityState.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Preserved,
                    LegacyPptCapabilityState.Native,
                    "Document slide defaults, notes/handout defaults, main-master and per-slide overrides decode with all six flags, date format ids, and fixed strings. Native Open XML master/layout settings project back to binary PPT/POT/PPS. Imported per-slide flag and text edits append preservation-aware records; edits to shared binary master/default scopes remain loss-blocked."),
                new LegacyPptCapability(LegacyPptFeature.Comments, "Review",
                    "Classic comment authors and comment records.",
                    LegacyPptRepresentability.Native, LegacyPptCapabilityState.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Native,
                    LegacyPptCapabilityState.Native,
                    "PowerPoint 2000-2003 PP10 comments decode with author, initials, text, index, UTC creation time, and master-unit position. Classic comments project editably, write natively to PPT/POT/PPS, and imported additions, edits, and removals append preservation-aware slide records without replacing unrelated PP10 tag data."),
                Blocked(LegacyPptFeature.ModernComments, "Review",
                    "Modern threaded comments, replies, status, and shape anchors.",
                    "PowerPoint 97-2003 has no native threaded-comment model; conversion is explicitly loss-blocked."),
                Planned(LegacyPptFeature.Transitions, "Presentation", "Slide transitions, speed, and advance settings.",
                    "The complete SlideShowSlideInfoAtom is decoded, including every legacy effect id, direction, speed, sound reference, show flag, click advance, and timed advance. Cut, fade, wipe, blinds, comb, and directional push project editably and write natively with the three legacy speeds and advance timing. Remaining legacy effects, transition sound projection, and imported transition edits remain planned."),
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
                new LegacyPptCapability(LegacyPptFeature.AccessibilityMetadata, "Accessibility",
                    "Alternative text and object names.", LegacyPptRepresentability.Native,
                    LegacyPptCapabilityState.Native, LegacyPptCapabilityState.Planned,
                    LegacyPptCapabilityState.Preserved, LegacyPptCapabilityState.Planned,
                    "OfficeArt object names and descriptions are decoded and projected to native non-visual metadata for supported slide, master, layout, group, connector, and picture shapes. Unmodified and unrelated supported edits preserve the binary properties; editing this metadata, accessibility titles, decorative state, and fresh binary authoring remain planned."),
                new LegacyPptCapability(LegacyPptFeature.UnknownRecordsAndStreams, "Preservation",
                    "Unknown live records and compound streams.", LegacyPptRepresentability.Opaque,
                    LegacyPptCapabilityState.Preserved, LegacyPptCapabilityState.Blocked,
                    LegacyPptCapabilityState.Planned, LegacyPptCapabilityState.Blocked,
                    "No-op saves retain the exact package. Supported mapped shape and text edits append a UserEdit and preserve untouched records and streams; broader edited preservation remains planned.")
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
