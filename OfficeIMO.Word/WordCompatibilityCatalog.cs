using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

/// <summary>Versioned capability contract for Word 97-2003 DOC interoperability.</summary>
public static class WordCompatibilityCatalog {
    private const string FormatId = "Word.Doc";

    /// <summary>Gets the current machine-readable DOC compatibility contract.</summary>
    public static OfficeCapabilityCatalog Current { get; } = new(
        "OfficeIMO.Word.LegacyDoc",
        schemaVersion: 1,
        new[] {
            Native("Lifecycle.File", "Lifecycle", "Path, stream, byte, synchronous, and asynchronous document flows."),
            Native("Lifecycle.Variants", "Lifecycle", "DOC document and DOT binary-template routing, import, native writing, and template FIB semantics."),
            Native("Structure.MainStory", "Structure", "Main-story paragraphs, runs, tabs, and line, page, and column breaks."),
            Native("Structure.Sections", "Structure", "Sections, page size, margins, orientation, columns, and supported section flags."),
            Subset("Structure.HeadersFooters", "Structure", "Default, first-page, and even-page headers and footers.",
                "Text, fields, links, bookmarks, content controls, and inline pictures are supported within the documented story subset."),
            Subset("Text.CharacterFormatting", "Text", "Character formatting, fonts, colors, language, and text effects supported by DOC.",
                "Common and tested sprm mappings are editable; unsupported Open XML-only effects block legacy output."),
            Subset("Text.ParagraphFormatting", "Text", "Paragraph styles, alignment, spacing, indents, borders, shading, tabs, and pagination.",
                "Common and tested paragraph properties map natively; unsupported properties are preflight-blocked."),
            Subset("Text.Numbering", "Text", "Bullets, numbering, list levels, and list overrides.",
                "The supported classic numbering subset maps editably in both directions."),
            Subset("Styles.ParagraphCharacter", "Styles", "Built-in and custom paragraph and character styles.",
                "Supported style inheritance and formatting are native; unsupported style properties are diagnosed."),
            Subset("Styles.Table", "Styles", "Table styles and conditional table formatting.",
                "The BIFF-era Word table-style subset is encoded; Open XML-only conditional properties are blocked."),
            Subset("Tables.Content", "Tables", "Tables, rows, cells, merges, widths, borders, shading, and nested tables.",
                "Supported tables are editable; native writing currently limits nested-table depth and rejects unsupported shapes."),
            Native("Navigation.Bookmarks", "Navigation", "Bookmarks in document stories and supported table boundaries."),
            Subset("Navigation.Hyperlinks", "Navigation", "Internal bookmark and absolute external hyperlinks.",
                "Simple supported display content is native; unsupported hyperlink payloads block output."),
            Subset("Fields.Common", "Fields", "Common simple and complex field instructions and display results.",
                "Supported fields retain instructions and results; unknown or active fields require an explicit fallback decision."),
            Native("Notes.Footnotes", "Notes", "Footnote bodies, references, separators, and supported formatting."),
            Native("Notes.Endnotes", "Notes", "Endnote bodies, references, separators, and supported formatting."),
            Subset("Review.Comments", "Review", "Classic comments, authors, ranges, and comment text.",
                "Readable classic comments are projected. Rich or structurally unsupported comment content remains guarded."),
            Subset("Review.Revisions", "Review", "Tracked insertions, deletions, authors, dates, and revision settings.",
                "Supported run-level revisions write natively; moves, property revisions, and unsupported nested markup block legacy output."),
            Subset("Controls.InlineContentControls", "Controls", "Inline content controls containing supported run-level content.",
                "DOC has no Open XML content-control carrier; supported content is converted to equivalent editable story content."),
            NativeWithVisualFallback("Drawing.InlinePictures", "Drawing", "Inline raster and metafile pictures.",
                "Supported encodings write as editable inline DOC pictures. If the native writer rejects a modern picture or effect, PreferVisual and BestEffort create static page images; PreservationOnly can also retain the source carrier."),
            VisualFallback("Drawing.FloatingObjects", "Drawing", "Floating pictures, text boxes, AutoShapes, groups, and drawing anchors.",
                "Binary OfficeArt is detected or preserved. Modern-to-DOC conversion either blocks or creates deterministic static page images after an omission-free render; PreservationOnly can also retain the source carrier."),
            VisualFallback("Drawing.Charts", "Drawing", "Embedded Word charts and chart previews.",
                "Modern charts render through the shared Drawing chart model. Strict conversion blocks; visual modes create static page images, and PreservationOnly can also retain the source carrier."),
            VisualFallback("Drawing.SmartArt", "Drawing", "SmartArt diagrams and their visible drawing fallback.",
                "DOC has no SmartArt model. Strict conversion blocks; visual modes render the diagram into static page images, and PreservationOnly can also retain the source carrier."),
            Opaque("Embedded.Ole", "Embedded", "Embedded OLE packages and their preview images.",
                OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Carrier),
            Opaque("Embedded.ActiveX", "Embedded", "ActiveX controls and control storages.",
                OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier),
            Opaque("Embedded.Vba", "Embedded", "VBA project storage and related macro state.",
                OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier),
            Native("Metadata.BuiltIn", "Metadata", "Built-in SummaryInformation and DocumentSummaryInformation properties."),
            Native("Metadata.Custom", "Metadata", "Supported scalar custom document properties."),
            ModernOnly("Metadata.CustomXml", "Metadata", "Open XML custom XML parts.",
                "Classic DOC has no package-part equivalent; strict conversion blocks and permissive conversion must report the carrier loss."),
            Opaque("Security.DigitalSignatures", "Security", "Legacy digital-signature streams and storages.",
                OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier),
            new OfficeCapability(
                "Word.Security.Encryption", FormatId, OfficeDocumentFamily.Word, "Security",
                "Legacy DOC password-to-open encryption.", OfficeCapabilityRepresentability.Native,
                OfficeCapabilityCoverageState.NotImplemented,
                OfficeCapabilityCoverageState.NotImplemented,
                OfficeCapabilityCoverageState.PreservedOpaque,
                OfficeCapabilityCoverageState.Blocked,
                OfficeCapabilityCoverageState.Blocked,
                OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier,
                "Encrypted DOC detection is safe, but first-party decryption and encryption are not yet complete."),
            Opaque("Preservation.UnknownRecords", "Preservation", "Unknown records, table-stream ranges, compound streams, and storages.",
                OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability),
            new OfficeCapability(
                "Word.Preservation.SourceCarrier", FormatId, OfficeDocumentFamily.Word, "Preservation",
                "Hash-verified original source retained beside an allowed lossy conversion or visual fallback.", OfficeCapabilityRepresentability.Opaque,
                OfficeCapabilityCoverageState.PreservedOpaque,
                OfficeCapabilityCoverageState.NotApplicable,
                OfficeCapabilityCoverageState.PreservedOpaque,
                OfficeCapabilityCoverageState.EmbeddedSource,
                OfficeCapabilityCoverageState.EmbeddedSource,
                OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Security,
                "Embedding is explicit because the original bytes may contain macros or hidden content. PreservationOnly enables it automatically; callers can recover the verified payload through TryGetCompatibilitySourcePayload.")
        });

    private static OfficeCapability Native(string id, string category, string description) => new(
        "Word." + id, FormatId, OfficeDocumentFamily.Word, category, description,
        OfficeCapabilityRepresentability.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Native,
        OfficeCapabilityCoverageState.Native);

    private static OfficeCapability Subset(string id, string category, string description, string note) => new(
        "Word." + id, FormatId, OfficeDocumentFamily.Word, category, description,
        OfficeCapabilityRepresentability.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Approximated,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Approximated,
        OfficeCapabilityCoverageState.Native,
        OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Editability, note);

    private static OfficeCapability NativeWithVisualFallback(string id, string category, string description, string note) => new(
        "Word." + id, FormatId, OfficeDocumentFamily.Word, category, description,
        OfficeCapabilityRepresentability.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Native,
        OfficeCapabilityCoverageState.Native, OfficeCapabilityCoverageState.Rasterized,
        OfficeCapabilityCoverageState.Native,
        OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Carrier, note);

    private static OfficeCapability VisualFallback(string id, string category, string description, string note) => new(
        "Word." + id, FormatId, OfficeDocumentFamily.Word, category, description,
        OfficeCapabilityRepresentability.Approximation,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Rasterized,
        OfficeCapabilityCoverageState.Dropped,
        OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Carrier, note);

    private static OfficeCapability Opaque(
        string id,
        string category,
        string description,
        OfficeCompatibilityImpact impact) => new(
        "Word." + id, FormatId, OfficeDocumentFamily.Word, category, description,
        OfficeCapabilityRepresentability.Opaque,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.PreservedOpaque, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.Dropped, impact,
        "Opaque carrier retention is safe for unmodified legacy round trips; editable cross-format conversion is not claimed.");

    private static OfficeCapability ModernOnly(string id, string category, string description, string note) => new(
        "Word." + id, FormatId, OfficeDocumentFamily.Word, category, description,
        OfficeCapabilityRepresentability.NotRepresentable,
        OfficeCapabilityCoverageState.NotApplicable, OfficeCapabilityCoverageState.NotApplicable,
        OfficeCapabilityCoverageState.NotApplicable, OfficeCapabilityCoverageState.Blocked,
        OfficeCapabilityCoverageState.NotApplicable,
        OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability, note);
}
