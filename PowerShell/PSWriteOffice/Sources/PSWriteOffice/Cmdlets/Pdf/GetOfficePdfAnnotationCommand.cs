using System.Management.Automation;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets page annotations and action evidence from a PDF.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfAnnotation", DefaultParameterSetName = "Path")]
[OutputType(typeof(PdfAnnotation))]
public sealed class GetOfficePdfAnnotationCommand : PSCmdlet
{
    /// <summary>PDF input paths; supports wildcards.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    public string[] Path { get; set; } = System.Array.Empty<string>();

    /// <summary>Exact PDF paths, including literal brackets.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = System.Array.Empty<string>();

    /// <summary>Caller-owned live PDF; inspection never disposes it.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public PdfDocument? Document { get; set; }

    /// <summary>Bounded input load settings.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath"), OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }

    /// <summary>Optional annotation subtype filter such as Text, Link, Widget, or FreeText.</summary>
    [Parameter]
    public string? Subtype { get; set; }

    /// <summary>Optional one-based page number filter.</summary>
    [Parameter]
    public int? Page { get; set; }

    /// <summary>Return only annotations with actions, additional actions, or chained actions.</summary>
    [Parameter]
    public SwitchParameter WithAction { get; set; }

    /// <summary>Password used to inspect a Standard password-encrypted PDF.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == "Document")
        {
            WriteInspection(Document!);
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (var path in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal))
            WriteInspection(PdfDocument.Load(path, PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions)));
    }

    private void WriteInspection(PdfDocument document)
    {
        foreach (var annotation in document.Inspect().Annotations)
        {
            if (!string.IsNullOrWhiteSpace(Subtype) && !string.Equals(annotation.Subtype, Subtype, System.StringComparison.Ordinal))
            {
                continue;
            }

            if (Page.HasValue && annotation.PageNumber != Page.Value)
            {
                continue;
            }

            if (WithAction.IsPresent && !annotation.HasAction && !annotation.HasAdditionalActions && !annotation.HasChainedActions)
            {
                continue;
            }

            WriteObject(annotation);
        }
    }
}
