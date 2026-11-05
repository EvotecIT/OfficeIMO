using OfficeIMO.Excel;
using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Renders native pages, slides, sheets, ranges or charts and saves each image atomically.</summary>
/// <example>
/// <summary>Export selected slides.</summary>
/// <code>Export-OfficeDocumentImage -Path ./review.pptx -Page 1-3 -Format Svg -OutputDirectory ./slides</code>
/// </example>
/// <example>
/// <summary>Export exactly one worksheet range.</summary>
/// <code>Export-OfficeDocumentImage -LiteralPath './budget[final].xlsx' -Sheet Dashboard -Range A1:H30 -OutputPath ./dashboard.png -NoClobber</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeDocumentImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficeDocumentImageCommand : OfficeDocumentImageCmdlet {
    /// <summary>Input paths, with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathSheet")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathRange")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathChart")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathFile")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathFileSheet")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathFileRange")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathFileChart")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public override string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal paths; accepts FileInfo, PSPath and FullName properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathSheet")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathRange")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathChart")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathFile")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathFileSheet")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathFileRange")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathFileChart")]
    [Alias("PSPath", "FullName")]
    [ValidateNotNullOrEmpty]
    public override string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned native document, including current unsaved content.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentSheet")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentRange")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentChart")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentFile")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentFileSheet")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentFileRange")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentFileChart")]
    [ValidateNotNullOrEmpty]
    public override object Document { get; set; } = null!;

    /// <summary>Literal HTML markup.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Text")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "TextFile")]
    [AllowEmptyString]
    public override string Text { get; set; } = string.Empty;

    /// <summary>Markup source format; use Html.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Text")]
    [Parameter(Mandatory = true, ParameterSetName = "TextFile")]
    [ValidateNotNullOrEmpty]
    public override string From { get; set; } = string.Empty;

    /// <summary>Destination folder for the selected image batch.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Path")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "PathSheet")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "PathRange")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "PathChart")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPath")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPathSheet")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPathRange")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPathChart")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Document")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "DocumentSheet")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "DocumentRange")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "DocumentChart")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Text")]
    [ValidateNotNullOrEmpty]
    public string? OutputDirectory { get; set; }

    /// <summary>Exact single-image destination; its extension selects the default encoding.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "PathFile")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "PathFileSheet")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "PathFileRange")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "PathFileChart")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPathFile")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPathFileSheet")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPathFileRange")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPathFileChart")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "DocumentFile")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "DocumentFileSheet")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "DocumentFileRange")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "DocumentFileChart")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "TextFile")]
    [ValidateNotNullOrEmpty]
    public string? OutputPath { get; set; }

    /// <summary>Refuse to replace an existing destination image.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>One-based pages, slides or diagram pages. Cannot be combined with Excel selectors.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextFile")]
    [OfficePageRangeTransformation]
    public override OfficePageRange? Page { get; set; }

    /// <summary>Excel worksheets by name, one-based position or native sheet object.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathSheet")]
    [Parameter(ParameterSetName = "PathRange")]
    [Parameter(ParameterSetName = "PathChart")]
    [Parameter(Mandatory = true, ParameterSetName = "PathFileSheet")]
    [Parameter(ParameterSetName = "PathFileRange")]
    [Parameter(ParameterSetName = "PathFileChart")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathSheet")]
    [Parameter(ParameterSetName = "LiteralPathRange")]
    [Parameter(ParameterSetName = "LiteralPathChart")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathFileSheet")]
    [Parameter(ParameterSetName = "LiteralPathFileRange")]
    [Parameter(ParameterSetName = "LiteralPathFileChart")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentSheet")]
    [Parameter(ParameterSetName = "DocumentRange")]
    [Parameter(ParameterSetName = "DocumentChart")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentFileSheet")]
    [Parameter(ParameterSetName = "DocumentFileRange")]
    [Parameter(ParameterSetName = "DocumentFileChart")]
    [OfficeSheetTransformation]
    public override OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>An A1 range in the selected Excel sheets; excludes Chart and Page.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathRange")]
    [Parameter(Mandatory = true, ParameterSetName = "PathFileRange")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathRange")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathFileRange")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentRange")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentFileRange")]
    public override string? Range { get; set; }

    /// <summary>A chart name in the selected Excel sheets; excludes Range and Page.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathChart")]
    [Parameter(Mandatory = true, ParameterSetName = "PathFileChart")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathChart")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathFileChart")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentChart")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentFileChart")]
    public override string? Chart { get; set; }

    /// <inheritdoc />
    protected override bool WritesFiles => true;
    /// <inheritdoc />
    protected override string? ImageOutputPath => OutputPath;
    /// <inheritdoc />
    protected override string? ImageOutputDirectory => OutputDirectory;
    /// <inheritdoc />
    protected override bool PreventOverwrite => NoClobber.IsPresent;
}
