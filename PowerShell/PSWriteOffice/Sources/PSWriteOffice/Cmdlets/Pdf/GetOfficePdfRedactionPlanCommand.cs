using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Previews explicit or text-matched PDF redaction areas before content is removed.</summary>
/// <remarks>
/// This command reports redaction impact only. It does not remove or rewrite PDF content.
/// </remarks>
/// <example>
///   <summary>Preview a redaction rectangle.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePdfRedactionPlan -Path .\Report.pdf -PageNumber 1 -X 72 -Y 650 -Width 240 -Height 32</code>
///   <para>Returns line-level text blocks and annotations that intersect the rectangle.</para>
/// </example>
/// <example>
///   <summary>Find content to redact by text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePdfRedactionPlan -Path .\Report.pdf -Text 'Account number'</code>
///   <para>Derives reviewable redaction areas from logical text blocks containing the supplied text.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfRedactionPlan", DefaultParameterSetName = ParameterSetRectangle)]
[OutputType(typeof(PdfRedactionPlan))]
public sealed class GetOfficePdfRedactionPlanCommand : PSCmdlet
{
    private const string ParameterSetRectangle = "Rectangle";
    private const string ParameterSetArea = "Area";
    private const string ParameterSetText = "Text";

    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>One-based page number for the redaction rectangle.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public int PageNumber { get; set; }

    /// <summary>Left coordinate in PDF points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public double X { get; set; }

    /// <summary>Bottom coordinate in PDF points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public double Y { get; set; }

    /// <summary>Rectangle width in PDF points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public double Width { get; set; }

    /// <summary>Rectangle height in PDF points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public double Height { get; set; }

    /// <summary>Optional redaction area label.</summary>
    [Parameter(ParameterSetName = ParameterSetRectangle)]
    public string? Label { get; set; }

    /// <summary>One or more pre-created OfficeIMO.Pdf redaction areas.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetArea)]
    public PdfRedactionArea[] Area { get; set; } = System.Array.Empty<PdfRedactionArea>();

    /// <summary>Literal text used to derive redaction areas from matching logical text blocks.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string[] Text { get; set; } = System.Array.Empty<string>();

    /// <summary>Use case-sensitive literal text matching.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    public SwitchParameter MatchCase { get; set; }

    /// <summary>Password used to read a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfDocument.Load(
            PdfCommandUtilities.ResolvePath(this, Path),
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));

        if (ParameterSetName == ParameterSetText)
        {
            var search = new PdfRedactionSearchOptions { MatchCase = MatchCase.IsPresent };
            search.AddLiteral(Text);
            WriteObject(document.Redactions.Search(search));
            return;
        }

        var areas = ParameterSetName == ParameterSetArea
            ? Area
            : new[] { new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label) };

        WriteObject(document.Redactions.Plan(areas));
    }
}
