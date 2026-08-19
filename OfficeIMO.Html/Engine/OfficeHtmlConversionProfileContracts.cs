namespace OfficeIMO.Html;

/// <summary>
/// Shared catalog for source-specific Office-to-HTML conversion lane contracts.
/// </summary>
public static class OfficeHtmlConversionProfileContracts {
    private static readonly IReadOnlyList<OfficeHtmlConversionProfileContract> Contracts = new List<OfficeHtmlConversionProfileContract> {
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.WordSemanticDocument,
            "Word",
            "Word Semantic Document",
            HtmlConversionProfile.Semantic,
            "Word document headings, paragraphs, lists, tables, images, notes, comments, bookmarks, forms, and metadata as accessible HTML.",
            "Preserve document structure and readable content first; represent Word-only layout and review details as semantic regions, annotations, or diagnostics.",
            "none",
            new[] { "document sections", "headings", "paragraphs", "lists", "tables", "images", "figures", "bookmarks", "forms", "footnotes", "endnotes", "comments" },
            new[] { "shared asset manifest", "safe hyperlink policy", "embedded image data URI support", "optional metadata export" },
            new[] { "unsupported style diagnostics", "comment exposure opt-in boundary", "layout simplification diagnostics" }),
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.WordDocumentRoundTrip,
            "Word",
            "Word Document Roundtrip",
            HtmlConversionProfile.Document,
            "Editable HTML to Word to HTML scenarios where document semantics, forms, tables, resources, diagnostics, and package validity are reviewed together.",
            "Preserve editable document content and report lossy HTML or Word features without claiming browser layout or pixel-perfect print reconstruction.",
            "none",
            new[] { "HTML document shell", "head metadata", "headings", "paragraphs", "lists", "tables", "table sections", "form controls", "images", "notes", "comments", "headers-and-footers" },
            new[] { "shared resource manifest", "Word-owned image embedding", "trusted document resource policy", "roundtrip score" },
            new[] { "HTML import diagnostics", "OpenXML validation proof", "comment/reporting diagnostics", "unsupported CSS diagnostics" }),
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.WordPrintReview,
            "Word",
            "Word Print Review",
            HtmlConversionProfile.HighFidelityPrint,
            "Print-oriented Word HTML exports where section metadata, margins, headers, footers, tables, images, notes, and generated CSS are useful for review.",
            "Expose print-like document structure and styling metadata while keeping visual proof honest about browser and Word layout differences.",
            "OfficeIMO.Pdf",
            new[] { "section wrappers", "page setup metadata", "headers", "footers", "tables", "images", "figures", "notes", "comments", "style classes" },
            new[] { "shared asset manifest", "print media profile", "PDF owner for final print fidelity proof" },
            new[] { "browser-layout boundary", "print-fidelity boundary", "unsupported floating layout diagnostics" }),
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.ExcelSemanticTables,
            "Excel",
            "Excel Semantic Tables",
            HtmlConversionProfile.Semantic,
            "Workbook, worksheet, table, and cell data that should be readable, searchable, and accessible as HTML.",
            "Preserve workbook structure and values first; represent formulas, comments, and unsupported visual details as annotations or diagnostics.",
            "none",
            new[] { "workbook metadata", "worksheet sections", "tables", "row-and-column headers", "cell values", "formula annotations", "comments-as-notes" },
            new[] { "shared asset manifest", "safe hyperlink policy", "image/chart references reported but not visualized unless adapter opts in" },
            new[] { "merged-cell simplification diagnostics", "formula/display-value diagnostics", "unsupported visual feature diagnostics" }),
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.ExcelVisualReview,
            "Excel",
            "Excel Visual Review",
            HtmlConversionProfile.PositionedReview,
            "Worksheet ranges, pages, and review snapshots where layout and drawing fidelity matter more than editability.",
            "Preserve inspectable worksheet geometry using shared Drawing/SVG/raster primitives while reporting non-editable review boundaries.",
            "OfficeIMO.Drawing",
            new[] { "page wrappers", "positioned cell regions", "positioned images", "shapes", "charts", "comments", "headers-and-footers", "source anchors" },
            new[] { "shared asset manifest", "Drawing-owned image and shape rendering", "hyperlink frame reporting" },
            new[] { "visual simplification diagnostics", "unsupported drawing effect diagnostics", "no-editable-reconstruction boundary" }),
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.PowerPointSemanticSlides,
            "PowerPoint",
            "PowerPoint Semantic Slides",
            HtmlConversionProfile.Semantic,
            "Slide outlines, notes, tables, media references, alt text, and speaker content as accessible HTML.",
            "Preserve slide reading order and semantic content first; represent complex layout as source anchors and diagnostics.",
            "none",
            new[] { "presentation metadata", "slide sections", "headings", "paragraphs", "lists", "tables", "images", "charts", "speaker notes", "alt text" },
            new[] { "shared asset manifest", "safe hyperlink policy", "media reference inventory" },
            new[] { "reading-order diagnostics", "unsupported SmartArt/media diagnostics", "layout simplification diagnostics" }),
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.PowerPointVisualReview,
            "PowerPoint",
            "PowerPoint Visual Review",
            HtmlConversionProfile.PositionedReview,
            "Slide previews, review pages, and proof galleries where slide appearance needs to remain inspectable in HTML.",
            "Preserve inspectable slide geometry using shared Drawing/SVG/raster primitives while avoiding editable reconstruction claims.",
            "OfficeIMO.Drawing",
            new[] { "slide wrappers", "positioned text frames", "positioned images", "shapes", "charts", "tables", "media placeholders", "source anchors" },
            new[] { "shared asset manifest", "Drawing-owned slide rendering", "media placeholder policy", "hyperlink frame reporting" },
            new[] { "visual simplification diagnostics", "unsupported animation/media diagnostics", "no-editable-reconstruction boundary" }),
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.RtfSemanticDocument,
            "RTF",
            "RTF Semantic Document",
            HtmlConversionProfile.Semantic,
            "RTF paragraphs, lists, tables, fields, links, images, notes, and safe object fallbacks as accessible HTML.",
            "Preserve readable document structure first while rejecting unsafe resources and reporting controls that require simplification or omission.",
            "none",
            new[] { "headings", "paragraphs", "lists", "tables", "fields", "links", "images", "notes", "safe object fallbacks" },
            new[] { "restrictive web URL policy", "bounded embedded image policy", "encapsulated HTML reparsing" },
            new[] { "per-construct preserve/simplify/omit actions", "resource rejection diagnostics", "producer-control preservation diagnostics" }),
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.RtfDocumentRoundTrip,
            "RTF",
            "RTF Document Roundtrip",
            HtmlConversionProfile.Document,
            "Trusted RTF-to-HTML-to-RTF workflows where editable structure, private metadata, binary resources, and preservation diagnostics are reviewed together.",
            "Preserve supported RTF controls and payloads through typed private metadata without presenting that metadata as safe public HTML.",
            "none",
            new[] { "document shell", "paragraphs", "lists", "nested tables", "fields", "images", "objects", "shapes", "notes", "headers-and-footers" },
            new[] { "OfficeIMO trusted URL policy", "embedded image payloads", "private roundtrip metadata" },
            new[] { "per-construct preserve/simplify/omit actions", "roundtrip preservation diagnostics", "unsafe-publication boundary" }),
        new OfficeHtmlConversionProfileContract(
            OfficeHtmlConversionProfile.RtfPrintReview,
            "RTF",
            "RTF Print Review",
            HtmlConversionProfile.HighFidelityPrint,
            "Styled browser and print review of RTF sections, tables, fields, images, notes, and page-break behavior.",
            "Expose a polished static document shell and print rules while retaining explicit RTF-to-browser layout boundaries.",
            "OfficeIMO.Pdf",
            new[] { "document shell", "section regions", "headings", "paragraphs", "lists", "tables", "fields", "images", "notes", "page breaks" },
            new[] { "shared OfficeIMO theme", "print media profile", "bounded web resource policy" },
            new[] { "per-construct preserve/simplify/omit actions", "browser-layout boundary", "print-fidelity boundary" })
    }.AsReadOnly();

    /// <summary>Gets all source-specific Office-to-HTML profile contracts.</summary>
    public static IReadOnlyList<OfficeHtmlConversionProfileContract> All => Contracts;

    /// <summary>Gets a source-specific Office-to-HTML profile contract by profile identifier.</summary>
    public static OfficeHtmlConversionProfileContract Get(OfficeHtmlConversionProfile profile) {
        foreach (OfficeHtmlConversionProfileContract contract in Contracts) {
            if (contract.Profile == profile) {
                return contract;
            }
        }

        throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unknown Office HTML conversion profile.");
    }
}
