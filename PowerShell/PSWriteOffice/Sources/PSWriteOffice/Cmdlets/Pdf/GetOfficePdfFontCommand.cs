using System;
using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Lists fonts with embedding and ToUnicode diagnostics.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfFont", DefaultParameterSetName = "Path")]
[OutputType(typeof(PdfFontDiagnostic))]
public sealed class GetOfficePdfFontCommand : PSCmdlet {
    /// <summary>Path for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Subtype for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public string? Subtype { get; set; }
    /// <summary>NeedsReview for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter NeedsReview { get; set; }
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
        foreach (var font in document.Diagnostics().Fonts) {
            if (Subtype != null && !string.Equals(font.Subtype, Subtype, StringComparison.Ordinal)) continue;
            if (NeedsReview && !font.RequiresEmbeddingReview && !font.RequiresToUnicodeReview) continue;
            WriteObject(font);
        }
    }
}
