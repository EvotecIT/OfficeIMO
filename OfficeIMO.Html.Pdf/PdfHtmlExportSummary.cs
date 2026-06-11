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
        int linkCount,
        int formFieldCount,
        int formWidgetCount,
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
        LinkCount = linkCount;
        FormFieldCount = formFieldCount;
        FormWidgetCount = formWidgetCount;
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

    /// <summary>Number of logical link annotations represented by the selected logical pages.</summary>
    public int LinkCount { get; }

    /// <summary>Number of document form fields available to the export.</summary>
    public int FormFieldCount { get; }

    /// <summary>Number of form widget placements represented by the selected logical pages.</summary>
    public int FormWidgetCount { get; }

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
