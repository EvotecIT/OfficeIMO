using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

public sealed partial class SetOfficePdfAnnotationCommand {
    /// <summary>Native annotation edit policy and fields. Explicit parameters override this captured policy; append-only residual data remains opt-in.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(PdfAnnotationUpdateOptions))]
    public PdfAnnotationUpdateOptions? Options { get; set; }

    /// <summary>Input PDF paths with wildcard expansion.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathUri", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathUriDirectory", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Single-file destination.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true)]
    [Parameter(ParameterSetName = "PathUri", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathUri", Mandatory = true)]
    public string? OutputPath { get; set; }
    /// <summary>ObjectNumber setting.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    public int ObjectNumber { get; set; }
    /// <summary>Contents setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentUri")]
    public string? Contents { get; set; }
    /// <summary>Title setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentUri")]
    public string? Title { get; set; }
    /// <summary>Name setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentUri")]
    public string? Name { get; set; }
    /// <summary>Raw PDF annotation flags.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentUri")]
    public int? Flags { get; set; }
    /// <summary>Replacement annotation RGB color.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentUri")]
    [OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }
    /// <summary>RemoveAction setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentUri")]
    public SwitchParameter RemoveAction { get; set; }
    /// <summary>Password setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>IgnorePermissionRestrictions setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>Native parsing, authentication and resource policy.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>Return native annotation edit reports.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathUri")]
    [Parameter(ParameterSetName = "PathUriDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathUri")]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentUri")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>Batch destination retaining source names.</summary>
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PathUriDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory", Mandatory = true)]
    public string? OutputDirectory { get; set; }
    /// <summary>Exact or wildcard target URI used to select link annotations.</summary>
    [Parameter(ParameterSetName = "PathUri", Mandatory = true)]
    [Parameter(ParameterSetName = "PathUriDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathUri", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "DocumentUri", Mandatory = true)]
    public string? Uri { get; set; }
    /// <summary>Literal input PDF paths; accepts FileInfo.FullName.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathUri", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathUriDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "PSPath")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Native caller-owned PDF; omission uses the enclosing content block.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "DocumentUri", ValueFromPipeline = true)]
    public PdfDocument? Document { get; set; }
}
