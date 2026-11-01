using System;
using System.Management.Automation;
using OfficeIMO.Drawing;

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
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathFile")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public override string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal paths; accepts FileInfo, PSPath and FullName properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathFile")]
    [Alias("PSPath", "FullName")]
    [ValidateNotNullOrEmpty]
    public override string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned native document, including current unsaved content.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentFile")]
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
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPath")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Document")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Text")]
    [ValidateNotNullOrEmpty]
    public string? OutputDirectory { get; set; }

    /// <summary>Exact single-image destination; its extension selects the default encoding.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "PathFile")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPathFile")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "DocumentFile")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "TextFile")]
    [ValidateNotNullOrEmpty]
    public string? OutputPath { get; set; }

    /// <summary>Refuse to replace an existing destination image.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <inheritdoc />
    protected override bool WritesFiles => true;
    /// <inheritdoc />
    protected override string? ImageOutputPath => OutputPath;
    /// <inheritdoc />
    protected override string? ImageOutputDirectory => OutputDirectory;
    /// <inheritdoc />
    protected override bool PreventOverwrite => NoClobber.IsPresent;
}
