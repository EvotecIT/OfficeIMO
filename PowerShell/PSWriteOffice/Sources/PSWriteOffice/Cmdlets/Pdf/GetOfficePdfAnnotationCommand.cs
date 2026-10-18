using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets page annotations and action evidence from a PDF.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfAnnotation")]
[OutputType(typeof(PdfAnnotation))]
public sealed class GetOfficePdfAnnotationCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional annotation subtype filter such as Text, Link, Widget, or FreeText.</summary>
    [Parameter]
    public string? Subtype { get; set; }

    /// <summary>Optional one-based page number filter.</summary>
    [Parameter]
    public int? PageNumber { get; set; }

    /// <summary>Return only annotations with actions, additional actions, or chained actions.</summary>
    [Parameter]
    public SwitchParameter WithAction { get; set; }

    /// <summary>Password used to inspect a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        foreach (var annotation in PdfDocument
                     .Load(
                         PdfCommandUtilities.ResolvePath(this, Path),
                         PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
                     .Inspect()
                     .Annotations)
        {
            if (!string.IsNullOrWhiteSpace(Subtype) && !string.Equals(annotation.Subtype, Subtype, System.StringComparison.Ordinal))
            {
                continue;
            }

            if (PageNumber.HasValue && annotation.PageNumber != PageNumber.Value)
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
