using System;
using System.Management.Automation;
using OfficeIMO.Drawing;

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
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public override string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal paths; accepts FileInfo, PSPath and FullName properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    [ValidateNotNullOrEmpty]
    public override string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned native document, including current unsaved content.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
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

    /// <inheritdoc />
    protected override bool WritesFiles => false;
}
