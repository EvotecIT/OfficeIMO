using System;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

public sealed partial class ImportOfficePdfXfdfCommand {
    /// <summary>Input PDF paths with wildcard expansion.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathFile", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathFileDirectory", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Single-file destination.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "PathFile", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathFile", Mandatory = true)]
    public string? OutputPath { get; set; }
    /// <summary>Xfdf setting.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    public string? Xfdf { get; set; }
    /// <summary>Options setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "PathFileDirectory")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "LiteralPathFileDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentFile")]
    public object[]? Options { get; set; } = Array.Empty<object>();
    /// <summary>MaxXfdfSize setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "PathFileDirectory")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "LiteralPathFileDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentFile")]
    public long MaxXfdfSize { get; set; } = 4L * 1024 * 1024;
    /// <summary>Password setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "PathFileDirectory")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "LiteralPathFileDirectory")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>IgnorePermissionRestrictions setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "PathFileDirectory")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "LiteralPathFileDirectory")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>Native parsing, authentication and resource policy.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "PathFileDirectory")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "LiteralPathFileDirectory")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>Return native annotation edit reports.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "PathFileDirectory")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "LiteralPathFileDirectory")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentFile")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>Batch destination retaining source names.</summary>
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PathFileDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathFileDirectory", Mandatory = true)]
    public string? OutputDirectory { get; set; }
    /// <summary>Literal input PDF paths; accepts FileInfo.FullName.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathFile", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathFileDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "PSPath")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>XfdfPath setting.</summary>
    [Parameter(ParameterSetName = "PathFile", Mandatory = true)]
    [Parameter(ParameterSetName = "PathFileDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathFile", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathFileDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "DocumentFile", Mandatory = true)]
    public string? XfdfPath { get; set; }
    /// <summary>Native caller-owned PDF; omission uses the enclosing content block.</summary>
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentFile", ValueFromPipeline = true)]
    public PdfDocument? Document { get; set; }
}
