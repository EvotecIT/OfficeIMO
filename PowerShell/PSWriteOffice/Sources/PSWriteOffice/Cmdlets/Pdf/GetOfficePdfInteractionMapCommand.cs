using System;
using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Builds text-selection and interactive hit regions for one page.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfInteractionMap", DefaultParameterSetName = "Path")]
[OutputType(typeof(PdfPageInteractionMap))]
public sealed class GetOfficePdfInteractionMapCommand : PSCmdlet {
    /// <summary>Path for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Page for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public int Page { get; set; } = 1;
    /// <summary>Options for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public object[]? Options { get; set; }
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
        var settings = new PdfPageInteractionOptions();
        var binder = new OfficeOptionsTransformationAttribute(typeof(PdfPageInteractionOptions));
        foreach (var option in Options ?? Array.Empty<object>()) {
            var value = OfficeArgumentValue.Unwrap(option);
            if (value is PdfPageInteractionOptions typed) settings = new PdfPageInteractionOptions { MaxTextRegions = typed.MaxTextRegions, MaxImageRegions = typed.MaxImageRegions, IncludeInvisibleText = typed.IncludeInvisibleText, IncludeTextRenderingMode3 = typed.IncludeTextRenderingMode3 };
            else if (value is System.Collections.IDictionary dictionary) settings = (PdfPageInteractionOptions)binder.ApplyTo((EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext"), settings, dictionary);
            else throw new PSArgumentException("Options accepts native PDF interaction settings and property overrides.");
        }
        WriteObject(PdfPageInteractionMap.Create(document.ToBytes(), Page, settings));
    }
}
