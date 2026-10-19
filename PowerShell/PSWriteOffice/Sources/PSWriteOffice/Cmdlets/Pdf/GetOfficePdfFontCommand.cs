using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets PDF font diagnostics for embedding and ToUnicode repair-readiness workflows.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfFont")]
[OutputType(typeof(PdfFontDiagnostic))]
public sealed class GetOfficePdfFontCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional font subtype filter such as Type1, Type0, or TrueType.</summary>
    [Parameter]
    public string? Subtype { get; set; }

    /// <summary>Return only fonts that need embedding or ToUnicode review.</summary>
    [Parameter]
    public SwitchParameter NeedsReview { get; set; }

    /// <summary>Password used to analyze a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        foreach (var font in PdfDocument
                     .Load(
                         PdfCommandUtilities.ResolvePath(this, Path),
                         PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
                     .Diagnostics()
                     .Fonts)
        {
            if (!string.IsNullOrWhiteSpace(Subtype) && !string.Equals(font.Subtype, Subtype, System.StringComparison.Ordinal))
            {
                continue;
            }

            if (NeedsReview.IsPresent && !font.RequiresEmbeddingReview && !font.RequiresToUnicodeReview)
            {
                continue;
            }

            WriteObject(font);
        }
    }
}
