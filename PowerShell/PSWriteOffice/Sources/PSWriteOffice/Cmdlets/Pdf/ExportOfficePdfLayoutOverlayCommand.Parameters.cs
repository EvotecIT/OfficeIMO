using System.Management.Automation;
using System.Security;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

public sealed partial class ExportOfficePdfLayoutOverlayCommand {
    /// <summary>Input PDF paths with wildcard expansion.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Single-file destination.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true)]
    public string? OutputPath { get; set; }
    /// <summary>Page setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public int Page { get; set; } = 1;
    /// <summary>Format setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Svg;
    /// <summary>Scale setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public double Scale { get; set; } = 1.0;
    /// <summary>Options setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public object[]? Options { get; set; } = Array.Empty<object>();
    /// <summary>LayoutOptions setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public PdfTextLayoutOptions? LayoutOptions { get; set; }
    /// <summary>NoClobber setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter NoClobber { get; set; }
    /// <summary>Password setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>IgnorePermissionRestrictions setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>Native parsing, authentication and resource policy.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>Return native annotation edit reports.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>Batch destination retaining source names.</summary>
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true)]
    public string? OutputDirectory { get; set; }
    /// <summary>Literal input PDF paths; accepts FileInfo.FullName.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "PSPath")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
}
