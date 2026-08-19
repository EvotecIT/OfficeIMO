namespace OfficeIMO.Html;

/// <summary>Single built-in capability and public-surface registry for HTML conversion targets.</summary>
public static class HtmlTargetCapabilityContracts {
    private static readonly HtmlSemanticFeature[] AllFeatures =
        global::OfficeIMO.Internal.EnumCompat.GetValues<HtmlSemanticFeature>();

    private static readonly IReadOnlyList<HtmlTargetCapabilityContract> Contracts = Array.AsReadOnly(new[] {
        Create(
            HtmlConversionTarget.Word,
            "OfficeIMO.Word.Html",
            "WordDocument",
            "HtmlConversionDocument.ToWordDocument",
            "HtmlToWordResult",
            "WordDocument.ToHtml",
            "HtmlTextConversionResult",
            "Load or LoadAsync the shared document, then use synchronous or asynchronous Word import APIs.",
            "Return HTML in memory or save it synchronously or asynchronously to a path or caller-owned stream.",
            new[] { "OfficeIMO", "UntrustedHtml", "TrustedDocument" },
            new[] { "SemanticDocument", "DocumentRoundTrip", "PrintReview" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes, HtmlSemanticFeature.Comments,
                HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Css, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Media, HtmlSemanticFeature.Geometry, HtmlSemanticFeature.PagedLayout),
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Notes, HtmlSemanticFeature.Comments, HtmlSemanticFeature.Annotations,
                HtmlSemanticFeature.Css, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Media, HtmlSemanticFeature.Forms, HtmlSemanticFeature.Charts,
                HtmlSemanticFeature.Geometry, HtmlSemanticFeature.PagedLayout)),
        Create(
            HtmlConversionTarget.Excel,
            "OfficeIMO.Excel.Html",
            "ExcelDocument",
            "HtmlConversionDocument.ToExcelDocument",
            "HtmlToExcelResult",
            "ExcelDocument.ToHtml",
            "HtmlTextConversionResult",
            "Load or LoadAsync the shared document, then import it synchronously into the workbook model.",
            "Return HTML in memory or save it synchronously or asynchronously to a path or caller-owned stream.",
            new[] { "Semantic", "Auto", "Generic" },
            new[] { "SemanticTables", "VisualReview" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Tables,
                HtmlSemanticFeature.Images, HtmlSemanticFeature.Comments, HtmlSemanticFeature.Annotations,
                HtmlSemanticFeature.Formulas, HtmlSemanticFeature.Charts, HtmlSemanticFeature.Geometry,
                HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Headings, HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText,
                HtmlSemanticFeature.Links, HtmlSemanticFeature.Lists, HtmlSemanticFeature.Css),
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Tables,
                HtmlSemanticFeature.Images, HtmlSemanticFeature.Comments, HtmlSemanticFeature.Annotations,
                HtmlSemanticFeature.Formulas, HtmlSemanticFeature.Charts, HtmlSemanticFeature.Geometry,
                HtmlSemanticFeature.Css, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Headings, HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText,
                HtmlSemanticFeature.Links, HtmlSemanticFeature.Lists, HtmlSemanticFeature.PagedLayout)),
        Create(
            HtmlConversionTarget.PowerPoint,
            "OfficeIMO.PowerPoint.Html",
            "PowerPointPresentation",
            "HtmlConversionDocument.ToPowerPointPresentation",
            "HtmlToPowerPointResult",
            "PowerPointPresentation.ToHtml",
            "PowerPointToHtmlResult",
            "Load or LoadAsync the shared document, then import it synchronously into the presentation model.",
            "Return HTML in memory or save it synchronously or asynchronously to a path or caller-owned stream.",
            new[] { "Semantic", "Auto", "Generic" },
            new[] { "SemanticSlides", "VisualReview" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Notes, HtmlSemanticFeature.Charts, HtmlSemanticFeature.Geometry,
                HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links, HtmlSemanticFeature.Lists,
                HtmlSemanticFeature.Css),
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images, HtmlSemanticFeature.Media,
                HtmlSemanticFeature.Notes, HtmlSemanticFeature.Charts, HtmlSemanticFeature.Geometry,
                HtmlSemanticFeature.Css, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Lists, HtmlSemanticFeature.Comments, HtmlSemanticFeature.Annotations,
                HtmlSemanticFeature.PagedLayout)),
        Create(
            HtmlConversionTarget.OneNote,
            "OfficeIMO.OneNote.Html",
            "OneNoteSection / OneNoteNotebook",
            "HtmlConversionDocument.ToOneNoteSection",
            "HtmlToOneNoteSectionResult / HtmlToOneNoteNotebookResult",
            "OneNoteSection.ToHtmlDocument",
            "HtmlTextConversionResult",
            "Load or LoadAsync the shared document, then import it synchronously into a section or notebook.",
            "Return semantic or visual HTML in memory or save it synchronously or asynchronously to a path or caller-owned stream.",
            new[] { "GenericSemantic" },
            new[] { "SemanticHtml", "VisualHtml" },
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
            "HtmlConversionDocument",
            "Load or LoadAsync the shared document, then convert it synchronously to Markdown text or a Markdown document.",
            "Return the shared HTML document in memory or save it synchronously or asynchronously to a path or caller-owned stream.",
            new[] { "OfficeIMO", "GitHubFlavoredMarkdown", "CommonMark", "Portable" },
            new[] { "OfficeIMO", "GitHubFlavoredMarkdown", "CommonMark", "Portable" },
            Features(HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings, HtmlSemanticFeature.Paragraphs,
                HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links, HtmlSemanticFeature.Lists,
                HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images, HtmlSemanticFeature.Notes,
                HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Media, HtmlSemanticFeature.Forms,
                HtmlSemanticFeature.Comments, HtmlSemanticFeature.Css)),
        Create(
            HtmlConversionTarget.Rtf,
            "OfficeIMO.Html.Rtf",
            "RtfDocument",
            "HtmlConversionDocument.ToRtfDocument",
            "HtmlToRtfResult",
            "RtfDocument.ToHtml",
            "RtfToHtmlResult",
            "Load or LoadAsync the shared document, then convert it synchronously; RTF path and caller-owned stream saves support synchronous and asynchronous forms.",
            "Return HTML and its report in memory or save it synchronously or asynchronously to a path or caller-owned stream.",
            new[] { "OfficeIMO", "UntrustedHtml" },
            new[] { "SemanticDocument", "DocumentRoundTrip", "PrintReview" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes, HtmlSemanticFeature.Comments,
                HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Media, HtmlSemanticFeature.Geometry, HtmlSemanticFeature.Css,
                HtmlSemanticFeature.PagedLayout),
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes, HtmlSemanticFeature.Comments,
                HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Css, HtmlSemanticFeature.Resources,
                HtmlSemanticFeature.PagedLayout),
            Features(HtmlSemanticFeature.Media, HtmlSemanticFeature.Geometry)),
        Create(
            HtmlConversionTarget.Pdf,
            "OfficeIMO.Html.Pdf",
            "PdfDocument / PDF bytes",
            "HtmlConversionDocument.ToPdfDocument",
            "PdfDocumentConversionResult",
            "PdfHtmlConverterExtensions.ToHtml",
            "PdfHtmlConversionResult",
            "Resolve resources through the shared synchronous or asynchronous render pipeline and return the PDF document plus conversion evidence.",
            "Return review HTML and its report in memory or save it synchronously or asynchronously to a path or caller-owned stream.",
            new[] { "PagedPrint" },
            new[] { "Semantic", "PositionedReview" },
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Geometry, HtmlSemanticFeature.Css, HtmlSemanticFeature.Resources,
                HtmlSemanticFeature.PagedLayout),
            Features(HtmlSemanticFeature.Media, HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes,
                HtmlSemanticFeature.Comments, HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Formulas,
                HtmlSemanticFeature.Charts),
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Paragraphs,
                HtmlSemanticFeature.Links, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Forms, HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Geometry,
                HtmlSemanticFeature.Css, HtmlSemanticFeature.Resources, HtmlSemanticFeature.PagedLayout),
            Features(HtmlSemanticFeature.Headings, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Lists)),
        Create(
            HtmlConversionTarget.Image,
            "OfficeIMO.Html",
            "PNG / JPEG / TIFF / SVG / WebP",
            "HtmlConversionDocument.ToPng / ToSvg / ToJpeg / ToTiff / ToWebp",
            "OfficeImageExportResult",
            null,
            null,
            "Use the shared synchronous or asynchronous render pipeline for in-memory, path, stream, and paged fluent image outputs.",
            null,
            new[] { "ContinuousScreen", "PagedPrint" },
            null,
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
            "Registered Reader handlers accept path or caller-owned stream input with cancellation and asynchronous document reads.",
            null,
            new[] { "Default", "Portable", "UntrustedHtml", "Mhtml" },
            null,
            Features(HtmlSemanticFeature.Metadata, HtmlSemanticFeature.Sections, HtmlSemanticFeature.Headings,
                HtmlSemanticFeature.Paragraphs, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links,
                HtmlSemanticFeature.Lists, HtmlSemanticFeature.Tables, HtmlSemanticFeature.Images,
                HtmlSemanticFeature.Media, HtmlSemanticFeature.Forms, HtmlSemanticFeature.Notes,
                HtmlSemanticFeature.Resources),
            Features(HtmlSemanticFeature.Comments, HtmlSemanticFeature.Annotations, HtmlSemanticFeature.Formulas,
                HtmlSemanticFeature.Charts, HtmlSemanticFeature.Geometry, HtmlSemanticFeature.Css,
                HtmlSemanticFeature.PagedLayout))
    });

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
        string htmlToTargetIoAndAsyncBoundary,
        string? targetToHtmlIoAndAsyncBoundary,
        IEnumerable<string> htmlToTargetProfiles,
        IEnumerable<string>? targetToHtmlProfiles,
        HtmlSemanticFeature[] supported,
        HtmlSemanticFeature[] approximated,
        HtmlSemanticFeature[]? targetToHtmlSupported = null,
        HtmlSemanticFeature[]? targetToHtmlApproximated = null) {
        var assigned = new HashSet<HtmlSemanticFeature>(supported);
        assigned.UnionWith(approximated);
        HtmlSemanticFeature[] unsupported = AllFeatures.Where(feature => !assigned.Contains(feature)).ToArray();
        var htmlToTarget = new HtmlToTargetCapabilityContract(importEntryPoint, importResultContract,
            htmlToTargetIoAndAsyncBoundary,
            importResultContract + " exposes ordered structured diagnostics with native, simplification, omission, and error outcomes.",
            htmlToTargetProfiles, supported, approximated, unsupported);
        HtmlSemanticFeature[] reverseSupported = targetToHtmlSupported ?? supported;
        HtmlSemanticFeature[] reverseApproximated = targetToHtmlApproximated ?? approximated;
        var reverseAssigned = new HashSet<HtmlSemanticFeature>(reverseSupported);
        reverseAssigned.UnionWith(reverseApproximated);
        HtmlSemanticFeature[] reverseUnsupported = AllFeatures.Where(feature => !reverseAssigned.Contains(feature)).ToArray();
        TargetToHtmlCapabilityContract? targetToHtml = exportEntryPoint == null
            ? null
            : new TargetToHtmlCapabilityContract(exportEntryPoint, exportResultContract!, targetToHtmlIoAndAsyncBoundary!,
                exportResultContract + " exposes an immutable conversion report and per-construct fidelity outcomes.",
                targetToHtmlProfiles ?? Array.Empty<string>(), reverseSupported, reverseApproximated, reverseUnsupported);
        return new HtmlTargetCapabilityContract(target, packageName, artifactName, htmlToTarget, targetToHtml);
    }

    private static HtmlSemanticFeature[] Features(params HtmlSemanticFeature[] features) => features;
}
