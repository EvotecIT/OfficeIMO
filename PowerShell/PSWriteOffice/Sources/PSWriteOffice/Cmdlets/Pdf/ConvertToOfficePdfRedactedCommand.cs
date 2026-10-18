using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Applies explicit or text-matched PDF redactions and writes a new PDF.</summary>
/// <example>
///   <summary>Apply a redaction rectangle.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfRedacted -Path .\Report.pdf -OutputPath .\Report-Redacted.pdf -PageNumber 1 -X 72 -Y 650 -Width 240 -Height 32</code>
///   <para>Removes matching text objects and annotations in the rectangle, then paints a redaction mark.</para>
/// </example>
/// <example>
///   <summary>Redact logical text blocks containing a value.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfRedacted -Path .\Report.pdf -OutputPath .\Report-Redacted.pdf -Text 'Account number'</code>
///   <para>Finds matching logical text blocks, removes their content, and paints the derived areas.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfRedacted", DefaultParameterSetName = ParameterSetRectangle, SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ConvertToOfficePdfRedactedCommand : PSCmdlet
{
    private const string ParameterSetRectangle = "Rectangle";
    private const string ParameterSetArea = "Area";
    private const string ParameterSetText = "Text";

    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

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

    /// <summary>Redaction fill color. Named colors and hexadecimal values are accepted. Defaults to black.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FillColor { get; set; }

    /// <summary>Paint only areas that match text or annotations in the redaction plan.</summary>
    [Parameter]
    public SwitchParameter OnlyPaintMatches { get; set; }

    /// <summary>Password used to read a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed modification restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        string inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        string outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!ShouldProcess(outputPath, "Write redacted PDF"))
        {
            return;
        }

        var options = new PdfRedactionApplyOptions
        {
            FillColor = PdfCommandUtilities.ParseColor(FillColor) ?? PdfColor.Black,
            PaintUnmatchedAreas = !OnlyPaintMatches.IsPresent
        };

        var document = PdfDocument.Load(
            inputPath,
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));

        PdfDocument redacted;
        if (ParameterSetName == ParameterSetText)
        {
            var search = new PdfRedactionSearchOptions { MatchCase = MatchCase.IsPresent };
            search.AddLiteral(Text);
            var plan = document.Redactions.Search(search);
            if (!plan.HasMatches)
            {
                throw new PSInvalidOperationException(
                    "No PDF text blocks matched -Text. The output PDF was not written.");
            }

            redacted = document.Redactions.Apply(plan, options);
        }
        else
        {
            PdfRedactionArea[] areas = ParameterSetName == ParameterSetArea
                ? Area
                : new[] { new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label) };
            redacted = document.Redactions.Apply(areas, options);
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        redacted.Save(outputPath).RequireSuccess();
        WriteObject(new FileInfo(outputPath));
    }
}
