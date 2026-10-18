using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Reads a PDF into OfficeIMO.Pdf's canonical structured document model.</summary>
/// <example>
///   <summary>Read paragraphs and tables from a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$logical = Read-OfficePdf -Path .\Examples\Documents\Report.pdf -Profile Structured
/// foreach ($page in $logical.Pages) {
///     $page.Paragraphs | ForEach-Object { $_.Text }
///     $page.Tables | Select-Object @{ Name = 'Page'; Expression = { $page.PageNumber } }, @{ Name = 'Rows'; Expression = { $_.Rows.Count } }
/// }</code>
///   <para>Returns the canonical OfficeIMO.Pdf result, including typed pages, paragraphs, tables, images, and diagnostics.</para>
/// </example>
[Cmdlet(VerbsCommunications.Read, "OfficePdf", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PdfDocumentReadResult))]
public sealed class ReadOfficePdfCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>An existing OfficeIMO.Pdf document, such as output from Get-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument? Document { get; set; }

    /// <summary>Advanced structured-read settings. Friendly parameters override the corresponding setting when explicitly supplied.</summary>
    [Parameter]
    public PdfReadOptions? Options { get; set; }

    /// <summary>Semantic reconstruction profile. Structured is the default; Fast omits optional document-wide enrichment.</summary>
    [Parameter]
    public PdfReadProfile Profile { get; set; } = PdfReadProfile.Structured;

    /// <summary>Optional page ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Password used to open a Standard password-encrypted PDF.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        PdfReadOptions effectiveOptions = CreateEffectiveOptions();
        PdfDocument document = ParameterSetName == ParameterSetDocument
            ? Document ?? throw new PSArgumentNullException(nameof(Document))
            : PdfDocument.Load(
                PdfCommandUtilities.ResolvePath(this, Path),
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));

        WriteObject(document.Read(effectiveOptions));
    }

    private PdfReadOptions CreateEffectiveOptions()
    {
        PdfReadOptions source = Options?.Clone() ?? PdfReadOptions.Default;
        bool profileWasBound = MyInvocation.BoundParameters.ContainsKey(nameof(Profile));
        bool pageRangeWasBound = MyInvocation.BoundParameters.ContainsKey(nameof(PageRange));

        return new PdfReadOptions
        {
            Profile = profileWasBound ? Profile : source.Profile,
            PageSelection = pageRangeWasBound
                ? string.IsNullOrWhiteSpace(PageRange)
                    ? null
                    : PdfPageSelection.Parse(PageRange!)
                : source.PageSelection,
            LayoutOptions = source.LayoutOptions,
            Pipeline = source.Pipeline
        };
    }
}
