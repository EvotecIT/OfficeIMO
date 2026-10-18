using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets simple AcroForm fields from a PDF.</summary>
/// <example>
///   <summary>Inspect fields before filling a form.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePdfFormField -Path .\Examples\Documents\Request.pdf |
///     Select-Object Name, FieldType, Value
/// Set-OfficePdfForm -Path .\Examples\Documents\Request.pdf -OutputPath .\Examples\Documents\Request-Filled.pdf -Field @{ Requester = 'Ada Lovelace' }</code>
///   <para>Reads form field names so the fill hashtable can use the right keys.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfFormField")]
[OutputType(typeof(PdfFormField))]
public sealed class GetOfficePdfFormFieldCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional field name filter.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Password used to inspect a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var fields = PdfDocument
            .Load(
                PdfCommandUtilities.ResolvePath(this, Path),
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
            .Inspect()
            .FormFields;
        foreach (var field in fields)
        {
            if (string.IsNullOrWhiteSpace(Name) || string.Equals(field.Name, Name, System.StringComparison.Ordinal))
            {
                WriteObject(field);
            }
        }
    }
}
