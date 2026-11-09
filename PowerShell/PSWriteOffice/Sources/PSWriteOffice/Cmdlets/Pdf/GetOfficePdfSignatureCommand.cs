using System.Management.Automation;
using System.Security;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Reports signature fields, byte ranges, DocMDP, DSS/LTV evidence and preservation markers.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfSignature", DefaultParameterSetName = "Path")]
[OutputType(typeof(PdfSignatureValidationReport))]
public sealed class GetOfficePdfSignatureCommand : PSCmdlet {
    /// <summary>Path for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Password for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>IgnorePermissionRestrictions for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>LoadOptions for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>LiteralPath for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Document for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    public PdfDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => PdfInspectionSource.Visit(this, ParameterSetName, Path, LiteralPath, Document, LoadOptions, Password, IgnorePermissionRestrictions, Inspect);

    private void Inspect(PdfDocument document) {
        WriteObject(document.Security.ValidateSignatures());
    }
}
