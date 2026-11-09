using System.Management.Automation;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

public sealed partial class NewOfficePdfSignatureCommand {
    /// <summary>Input PDF paths with wildcard expansion.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Single-file destination.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true)]
    public string? OutputPath { get; set; }
    /// <summary>FieldName setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public string? FieldName { get; set; } = "Signature1";
    /// <summary>Filter setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public string? Filter { get; set; } = "Adobe.PPKLite";
    /// <summary>SubFilter setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public PdfExternalSignatureSubFilter SubFilter { get; set; } = PdfExternalSignatureSubFilter.DetachedCms;
    /// <summary>Name setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public string? Name { get; set; }
    /// <summary>Reason setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public string? Reason { get; set; }
    /// <summary>Location setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public string? Location { get; set; }
    /// <summary>ContactInfo setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public string? ContactInfo { get; set; }
    /// <summary>ReservedSize setting.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public int ReservedSize { get; set; } = 32768;
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
