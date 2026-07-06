using System.Collections.Generic;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Observable summary for one PDF to HTML export.
/// </summary>
public sealed class PdfHtmlExportSummary {
    internal PdfHtmlExportSummary(
        PdfHtmlProfile profile,
        string profileId,
        IReadOnlyList<int> pageNumbers,
        int sourcePageCount,
        int renderedPageCount,
        int textBlockCount,
        int headingCount,
        int listItemCount,
        int tableCount,
        int imageCount,
        int imagePlacementCount,
        int imagePlaceholderCount,
        int linkCount,
        int renderedLinkCount,
        int renderedSafeUriLinkCount,
        int renderedUnsafeUriLinkCount,
        int renderedInternalDestinationLinkCount,
        int skippedLinkCount,
        int outlineCount,
        int renderedOutlineCount,
        int formFieldCount,
        int formWidgetCount,
        bool hasAcroFormXfa,
        int acroFormXfaPacketCount,
        int acroFormXfaStreamCount,
        int acroFormXfaPayloadByteCount,
        bool hasOpenAction,
        bool hasCatalogActions,
        bool hasPageActions,
        bool hasAnnotationActions,
        bool hasActiveContent,
        int potentiallyUnsafeActionCount,
        int javaScriptActionCount,
        int launchActionCount,
        int submitFormActionCount,
        int importDataActionCount,
        int catalogActionCount,
        int pageActionCount,
        int selectedPageActionCount,
        int annotationActionCount,
        int selectedAnnotationActionCount,
        int warningCount,
        bool emitsDocumentShell,
        PdfHtmlImageExportMode imageExportMode,
        string fidelityContract,
        string unsupportedScope) {
        Profile = profile;
        ProfileId = profileId;
        PageNumbers = pageNumbers;
        SourcePageCount = sourcePageCount;
        RenderedPageCount = renderedPageCount;
        TextBlockCount = textBlockCount;
        HeadingCount = headingCount;
        ListItemCount = listItemCount;
        TableCount = tableCount;
        ImageCount = imageCount;
        ImagePlacementCount = imagePlacementCount;
        ImagePlaceholderCount = imagePlaceholderCount;
        LinkCount = linkCount;
        RenderedLinkCount = renderedLinkCount;
        RenderedSafeUriLinkCount = renderedSafeUriLinkCount;
        RenderedUnsafeUriLinkCount = renderedUnsafeUriLinkCount;
        RenderedInternalDestinationLinkCount = renderedInternalDestinationLinkCount;
        SkippedLinkCount = skippedLinkCount;
        OutlineCount = outlineCount;
        RenderedOutlineCount = renderedOutlineCount;
        FormFieldCount = formFieldCount;
        FormWidgetCount = formWidgetCount;
        HasAcroFormXfa = hasAcroFormXfa;
        AcroFormXfaPacketCount = acroFormXfaPacketCount;
        AcroFormXfaStreamCount = acroFormXfaStreamCount;
        AcroFormXfaPayloadByteCount = acroFormXfaPayloadByteCount;
        HasOpenAction = hasOpenAction;
        HasCatalogActions = hasCatalogActions;
        HasPageActions = hasPageActions;
        HasAnnotationActions = hasAnnotationActions;
        HasActiveContent = hasActiveContent;
        PotentiallyUnsafeActionCount = potentiallyUnsafeActionCount;
        JavaScriptActionCount = javaScriptActionCount;
        LaunchActionCount = launchActionCount;
        SubmitFormActionCount = submitFormActionCount;
        ImportDataActionCount = importDataActionCount;
        CatalogActionCount = catalogActionCount;
        PageActionCount = pageActionCount;
        SelectedPageActionCount = selectedPageActionCount;
        AnnotationActionCount = annotationActionCount;
        SelectedAnnotationActionCount = selectedAnnotationActionCount;
        WarningCount = warningCount;
        EmitsDocumentShell = emitsDocumentShell;
        ImageExportMode = imageExportMode;
        FidelityContract = fidelityContract;
        UnsupportedScope = unsupportedScope;
    }

    /// <summary>PDF to HTML profile used for the export.</summary>
    public PdfHtmlProfile Profile { get; }

    /// <summary>Stable profile identifier for manifests, wrappers, and sidecars.</summary>
    public string ProfileId { get; }

    /// <summary>One-based source page numbers rendered into the HTML output, preserving duplicate selections.</summary>
    public IReadOnlyList<int> PageNumbers { get; }

    /// <summary>Total page count in the loaded logical source document.</summary>
    public int SourcePageCount { get; }

    /// <summary>Number of page instances rendered into the HTML output.</summary>
    public int RenderedPageCount { get; }

    /// <summary>Number of line-level text blocks represented by the selected logical pages.</summary>
    public int TextBlockCount { get; }

    /// <summary>Number of heading blocks represented by the selected logical pages.</summary>
    public int HeadingCount { get; }

    /// <summary>Number of list items represented by the selected logical pages.</summary>
    public int ListItemCount { get; }

    /// <summary>Number of detected table regions represented by the selected logical pages.</summary>
    public int TableCount { get; }

    /// <summary>Number of logical image resources represented by the selected logical pages.</summary>
    public int ImageCount { get; }

    /// <summary>Number of image placements represented by the selected logical pages.</summary>
    public int ImagePlacementCount { get; }

    /// <summary>Number of image placeholder elements emitted by the selected profile and output policy.</summary>
    public int ImagePlaceholderCount { get; }

    /// <summary>Number of logical link annotations represented by the selected logical pages.</summary>
    public int LinkCount { get; }

    /// <summary>Number of link annotations emitted by the selected profile and output policy.</summary>
    public int RenderedLinkCount { get; }

    /// <summary>Number of emitted URI links that use a safe HTML link target.</summary>
    public int RenderedSafeUriLinkCount { get; }

    /// <summary>Number of emitted URI links preserved as inert unsafe link metadata.</summary>
    public int RenderedUnsafeUriLinkCount { get; }

    /// <summary>Number of emitted internal destination links preserved as destination metadata.</summary>
    public int RenderedInternalDestinationLinkCount { get; }

    /// <summary>Number of selected logical links not emitted by the selected profile or output policy.</summary>
    public int SkippedLinkCount { get; }

    /// <summary>Number of outline/bookmark entries discovered in the loaded logical document.</summary>
    public int OutlineCount { get; }

    /// <summary>Number of outline/bookmark entries emitted into the selected HTML export scope.</summary>
    public int RenderedOutlineCount { get; }

    /// <summary>Number of document form fields available to the export.</summary>
    public int FormFieldCount { get; }

    /// <summary>Number of form widget placements represented by the selected logical pages.</summary>
    public int FormWidgetCount { get; }

    /// <summary>True when the source AcroForm contains XFA packets that are represented only as inert review metadata.</summary>
    public bool HasAcroFormXfa { get; }

    /// <summary>Number of named XFA packets discovered in the source AcroForm.</summary>
    public int AcroFormXfaPacketCount { get; }

    /// <summary>Number of XFA stream payloads discovered in the source AcroForm.</summary>
    public int AcroFormXfaStreamCount { get; }

    /// <summary>Total decoded byte length for readable XFA payloads.</summary>
    public int AcroFormXfaPayloadByteCount { get; }

    /// <summary>True when the exported page scope includes a readable document open action.</summary>
    public bool HasOpenAction { get; }

    /// <summary>True when catalog-level actions are represented in the exported document scope.</summary>
    public bool HasCatalogActions { get; }

    /// <summary>True when the exported pages include page-level actions.</summary>
    public bool HasPageActions { get; }

    /// <summary>True when the exported pages include annotation-level actions.</summary>
    public bool HasAnnotationActions { get; }

    /// <summary>True when catalog, page, or annotation actions are represented by the export summary.</summary>
    public bool HasActiveContent { get; }

    /// <summary>Number of represented actions whose type can execute script, launch external content, submit/import data, or play rich media.</summary>
    public int PotentiallyUnsafeActionCount { get; }

    /// <summary>Number of represented JavaScript actions.</summary>
    public int JavaScriptActionCount { get; }

    /// <summary>Number of represented launch actions.</summary>
    public int LaunchActionCount { get; }

    /// <summary>Number of represented form-submission actions.</summary>
    public int SubmitFormActionCount { get; }

    /// <summary>Number of represented import-data actions.</summary>
    public int ImportDataActionCount { get; }

    /// <summary>Number of catalog-level actions represented for the exported document scope.</summary>
    public int CatalogActionCount { get; }

    /// <summary>Number of page-level actions in the loaded source document.</summary>
    public int PageActionCount { get; }

    /// <summary>Number of page-level actions in the selected export page scope.</summary>
    public int SelectedPageActionCount { get; }

    /// <summary>Number of annotation-level actions in the loaded source document.</summary>
    public int AnnotationActionCount { get; }

    /// <summary>Number of annotation-level actions in the selected export page scope.</summary>
    public int SelectedAnnotationActionCount { get; }

    /// <summary>Number of conversion warnings recorded during the export.</summary>
    public int WarningCount { get; }

    /// <summary>True when the output includes a complete HTML document shell.</summary>
    public bool EmitsDocumentShell { get; }

    /// <summary>Image export behavior used for the generated HTML.</summary>
    public PdfHtmlImageExportMode ImageExportMode { get; }

    /// <summary>Profile fidelity contract copied into the export summary.</summary>
    public string FidelityContract { get; }

    /// <summary>Known unsupported or intentionally simplified scope copied into the export summary.</summary>
    public string UnsupportedScope { get; }
}
