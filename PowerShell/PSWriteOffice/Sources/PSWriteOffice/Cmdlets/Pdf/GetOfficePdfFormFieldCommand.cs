using System.Management.Automation;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
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
[Cmdlet(VerbsCommon.Get, "OfficePdfFormField", DefaultParameterSetName = "Path")]
[OutputType(typeof(PdfFormField))]
public sealed class GetOfficePdfFormFieldCommand : PSCmdlet
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

    /// <summary>Optional field name filter.</summary>
    [Parameter]
    public string? Name { get; set; }

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
        foreach (var field in document.Inspect().FormFields)
        {
            if (string.IsNullOrWhiteSpace(Name) || string.Equals(field.Name, Name, System.StringComparison.Ordinal))
            {
                WriteObject(field);
            }
        }
    }
}
