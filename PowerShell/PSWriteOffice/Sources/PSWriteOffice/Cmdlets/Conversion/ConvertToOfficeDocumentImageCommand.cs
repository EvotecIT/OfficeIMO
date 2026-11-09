using OfficeIMO.Excel;
using System.Management.Automation;
using OfficeIMO.Drawing;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Renders native pages, slides, sheets, ranges or charts to image results without writing files.</summary>
/// <example>
/// <summary>Render selected slides in memory.</summary>
/// <code>ConvertTo-OfficeDocumentImage -Path ./review.pptx -Page 1-3 -Format Svg | Select-Object Name, Width, Height</code>
/// </example>
/// <example>
/// <summary>Render HTML markup.</summary>
/// <code>ConvertTo-OfficeDocumentImage -Text '&lt;h1&gt;Service status&lt;/h1&gt;' -From Html</code>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeDocumentImage", DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ConvertToOfficeDocumentImageCommand : OfficeDocumentImageCmdlet {
    /// <summary>Input paths, with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathSheet")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathRange")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathChart")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public override string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal paths; accepts FileInfo, PSPath and FullName properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathSheet")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathRange")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathChart")]
    [Alias("PSPath", "FullName")]
    [ValidateNotNullOrEmpty]
    public override string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned native document, including current unsaved content.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentSheet")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentRange")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "DocumentChart")]
    [ValidateNotNullOrEmpty]
    public override object Document { get; set; } = null!;

    /// <summary>Literal HTML markup.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Text")]
    [AllowEmptyString]
    public override string Text { get; set; } = string.Empty;

    /// <summary>Markup source format; use Html.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Text")]
    [ValidateNotNullOrEmpty]
    public override string From { get; set; } = string.Empty;

    /// <summary>One-based pages, slides or diagram pages. Cannot be combined with Excel selectors.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Text")]
    [OfficePageRangeTransformation]
    public override OfficePageRange? Page { get; set; }

    /// <summary>Excel worksheets by name, one-based position or native sheet object.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathSheet")]
    [Parameter(ParameterSetName = "PathRange")]
    [Parameter(ParameterSetName = "PathChart")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathSheet")]
    [Parameter(ParameterSetName = "LiteralPathRange")]
    [Parameter(ParameterSetName = "LiteralPathChart")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentSheet")]
    [Parameter(ParameterSetName = "DocumentRange")]
    [Parameter(ParameterSetName = "DocumentChart")]
    [OfficeSheetTransformation]
    public override OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>An A1 range in the selected Excel sheets; excludes Chart and Page.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathRange")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathRange")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentRange")]
    public override string? Range { get; set; }

    /// <summary>A chart name in the selected Excel sheets; excludes Range and Page.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathChart")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathChart")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentChart")]
    public override string? Chart { get; set; }

    /// <inheritdoc />
    protected override bool WritesFiles => false;
}
