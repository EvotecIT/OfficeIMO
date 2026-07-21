namespace OfficeIMO.Html;

/// <summary>Single built-in capability and public-surface registry for HTML conversion targets.</summary>
public static class HtmlTargetCapabilityContracts {
    private static readonly HtmlSemanticFeature[] AllFeatures =
        (HtmlSemanticFeature[])Enum.GetValues(typeof(HtmlSemanticFeature));

    private static readonly IReadOnlyList<HtmlTargetCapabilityContract> Contracts = new[] {
        Create(
            HtmlConversionTarget.Word,
            "OfficeIMO.Word.Html",
            "WordDocument",
            "HtmlConversionDocument.ToWordDocument",
            "HtmlToWordResult",
            "WordDocument.ToHtml",
            "HtmlTextConversionResult",
            "Load or LoadAsync the shared document; synchronous and asynchronous Word import are available; path and stream HTML saves preserve caller-owned streams.",
            new[] { "OfficeIMO", "UntrustedHtml", "TrustedDocument" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes, HtmlSemanticFeature.Comments,
                HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Css, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Media, HtmlSemanticFeature.Geometry, HtmlSemanticFeature.PagedLayout)),
        Create(
            HtmlConversionTarget.Excel,
            "OfficeIMO.Excel.Html",
            "ExcelDocument",
            "HtmlConversionDocument.ToExcelDocument",
            "HtmlToExcelResult",
            "ExcelDocument.ToHtml",
            "HtmlTextConversionResult",
            "Load or LoadAsync the shared document, then import synchronously; path and stream HTML saves have synchronous and asynchronous forms.",
            new[] { "Semantic", "Auto", "Generic", "SemanticTables", "VisualReview" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Tables,
                HtmlSemanticFeature.Images, HtmlSemanticFeature.Comments, HtmlSemanticFeature.Annotations,
                HtmlSemanticFeature.Formulas, HtmlSemanticFeature.Charts, HtmlSemanticFeature.Geometry,
                HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Headings, HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText,
                HtmlSemanticFeature.Links, HtmlSemanticFeature.Lists, HtmlSemanticFeature.Css)),
        Create(
            HtmlConversionTarget.PowerPoint,
            "OfficeIMO.PowerPoint.Html",
            "PowerPointPresentation",
            "HtmlConversionDocument.ToPowerPointPresentation",
            "HtmlToPowerPointResult",
            "PowerPointPresentation.ToHtml",
            "PowerPointToHtmlResult",
            "Load or LoadAsync the shared document, then import synchronously; path and stream HTML saves have synchronous and asynchronous forms.",
            new[] { "Semantic", "Auto", "Generic", "SemanticSlides", "VisualReview" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Notes, HtmlSemanticFeature.Charts, HtmlSemanticFeature.Geometry,
                HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links, HtmlSemanticFeature.Lists,
                HtmlSemanticFeature.Css)),
        Create(
            HtmlConversionTarget.OneNote,
            "OfficeIMO.OneNote.Html",
            "OneNoteSection / OneNoteNotebook",
            "HtmlConversionDocument.ToOneNoteSection",
            "HtmlToOneNoteSectionResult / HtmlToOneNoteNotebookResult",
            "OneNoteSection.ToHtmlDocument",
            "HtmlTextConversionResult",
            "Load or LoadAsync the shared document, then import synchronously; semantic and visual HTML exports support path, stream, and asynchronous saves.",
            new[] { "GenericSemantic", "SemanticHtml", "VisualHtml" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Notes, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Geometry, HtmlSemanticFeature.Css)),
        Create(
            HtmlConversionTarget.Markdown,
            "OfficeIMO.Markdown.Html",
            "MarkdownDoc / Markdown text",
            "HtmlConversionDocument.ToMarkdownDocument",
            "HtmlToMarkdownResult",
            "MarkdownDoc.ToHtmlDocument",
            null,
            "Load or LoadAsync the shared document; Markdown conversion is synchronous and path or stream saves have synchronous and asynchronous forms.",
            new[] { "OfficeIMO", "GitHubFlavoredMarkdown", "CommonMark", "Portable" },
            Features(HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings, HtmlSemanticFeature.Paragraphs,
                HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links, HtmlSemanticFeature.Lists,
                HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images, HtmlSemanticFeature.Notes,
                HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Media, HtmlSemanticFeature.Forms,
                HtmlSemanticFeature.Comments, HtmlSemanticFeature.Css)),
        Create(
            HtmlConversionTarget.Rtf,
            "OfficeIMO.Html / OfficeIMO.Rtf",
            "RtfDocument",
            "HtmlConversionDocument.ToRtfDocument",
            "HtmlToRtfResult",
            "RtfDocument.ToHtml",
            "RtfToHtmlResult",
            "Load or LoadAsync the shared document; semantic conversion is synchronous and RTF/HTML path or stream saves have synchronous and asynchronous forms.",
            new[] { "OfficeIMO", "UntrustedHtml", "WebSafe", "RoundTrip" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes, HtmlSemanticFeature.Comments,
                HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Media, HtmlSemanticFeature.Geometry, HtmlSemanticFeature.Css,
                HtmlSemanticFeature.PagedLayout)),
        Create(
            HtmlConversionTarget.Pdf,
            "OfficeIMO.Html.Pdf",
            "PdfDocument / PDF bytes",
            "HtmlConversionDocument.ToPdfDocument",
            "PdfDocumentConversionResult",
            null,
            null,
            "Synchronous and asynchronous conversion resolve through the shared render resource pipeline; byte, document, path, and stream outputs are available.",
            new[] { "Continuous", "Paged", "Screen", "Print" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Geometry, HtmlSemanticFeature.Css, HtmlSemanticFeature.Resources,
                HtmlSemanticFeature.PagedLayout),
            Features(HtmlSemanticFeature.Media, HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes,
                HtmlSemanticFeature.Comments, HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Formulas,
                HtmlSemanticFeature.Charts)),
        Create(
            HtmlConversionTarget.Image,
            "OfficeIMO.Html",
            "PNG / JPEG / TIFF / SVG / WebP",
            "HtmlConversionDocument.ToPng / ToSvg / ToJpeg / ToTiff / ToWebp",
            "OfficeImageExportResult",
            null,
            null,
            "Synchronous and asynchronous render APIs share one resource pipeline; in-memory, path, stream, and paged fluent outputs are available.",
            new[] { "Continuous", "Paged", "Screen", "Print" },
            Features(HtmlSemanticFeature.Images, HtmlSemanticFeature.Geometry, HtmlSemanticFeature.Css,
                HtmlSemanticFeature.Resources, HtmlSemanticFeature.PagedLayout),
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Media,
                HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes, HtmlSemanticFeature.Comments,
                HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Formulas, HtmlSemanticFeature.Charts)),
        Create(
            HtmlConversionTarget.Reader,
            "OfficeIMO.Reader.Html",
            "OfficeDocumentReadResult / ReaderChunk",
            "OfficeDocumentReader.ReadDocument (after AddHtmlHandler)",
            "OfficeDocumentReadResult",
            null,
            null,
            "Registered Reader handlers support path and caller-owned stream input with cancellation and asynchronous document reads.",
            new[] { "Default", "Portable", "UntrustedHtml", "Mhtml" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Media, HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes,
                HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Comments, HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Formulas,
                HtmlSemanticFeature.Charts, HtmlSemanticFeature.Geometry, HtmlSemanticFeature.Css,
                HtmlSemanticFeature.PagedLayout))
    };

    /// <summary>Gets every built-in target contract in stable target order.</summary>
    public static IReadOnlyList<HtmlTargetCapabilityContract> All => Contracts;

    /// <summary>Gets one built-in target contract.</summary>
    public static HtmlTargetCapabilityContract Get(HtmlConversionTarget target) {
        foreach (HtmlTargetCapabilityContract contract in Contracts) {
            if (contract.Target == target) return contract;
        }

        throw new ArgumentOutOfRangeException(nameof(target), target, "Unknown HTML conversion target.");
    }

    private static HtmlTargetCapabilityContract Create(
        HtmlConversionTarget target,
        string packageName,
        string artifactName,
        string importEntryPoint,
        string importResultContract,
        string? exportEntryPoint,
        string? exportResultContract,
        string ioAndAsyncBoundary,
        IEnumerable<string> profiles,
        HtmlSemanticFeature[] supported,
        HtmlSemanticFeature[] approximated) {
        var assigned = new HashSet<HtmlSemanticFeature>(supported);
        assigned.UnionWith(approximated);
        HtmlSemanticFeature[] unsupported = AllFeatures.Where(feature => !assigned.Contains(feature)).ToArray();
        return new HtmlTargetCapabilityContract(target, packageName, artifactName, importEntryPoint,
            importResultContract, exportEntryPoint, exportResultContract, ioAndAsyncBoundary, profiles,
            supported, approximated, unsupported);
    }

    private static HtmlSemanticFeature[] Features(params HtmlSemanticFeature[] features) => features;
}
