using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Converts a PDF with simple AcroForm fields into a flat PDF.</summary>
/// <example>
///   <summary>Flatten a filled PDF form.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePdfForm -Path .\Examples\Documents\Request.pdf -OutputPath .\Examples\Documents\Request-Filled.pdf -Field @{
///     Requester = 'Ada Lovelace'
///     Priority = 'High'
/// }
/// ConvertTo-OfficePdfFlatForm -Path .\Examples\Documents\Request-Filled.pdf -OutputPath .\Examples\Documents\Request-Flat.pdf</code>
///   <para>Turns simple form fields into static page content.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfFlatForm", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ConvertToOfficePdfFlatFormCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed form-modification restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>TrueType or OpenType/CFF font file used to synthesize Unicode form field appearances while flattening.</summary>
    [Parameter]
    public string? AppearanceFontPath { get; set; }

    /// <summary>PDF font family name used for the supplied appearance font.</summary>
    [Parameter]
    public string? AppearanceFontFamilyName { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var formOptions = PdfCommandUtilities.CreateFormFillerOptions(this, AppearanceFontPath, AppearanceFontFamilyName, keepNeedAppearances: false);
        var document = PdfDocument.Load(
            PdfCommandUtilities.ResolvePath(this, Path),
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));
        var result = formOptions == null
            ? document.Forms.Flatten()
            : document.Forms.Flatten(formOptions);
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write flattened form PDF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        result.Save(outputPath).RequireSuccess();
        WriteObject(new FileInfo(outputPath));
    }
}
