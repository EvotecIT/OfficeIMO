using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Converts a PowerPoint deck to an HTML review document.</summary>
/// <example>
///   <summary>Export a deck as semantic HTML.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePowerPointHtml -Path .\Briefing.pptx -OutputPath .\Briefing.html -Title 'Briefing Review' -PassThru</code>
///   <para>Loads the deck and writes an HTML file with slide text, tables, pictures, charts, and notes where available.</para>
/// </example>
/// <example>
///   <summary>Export a visual review artifact.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePowerPointHtml -Path .\Briefing.pptx -Profile VisualReview -OutputPath .\Briefing.visual.html</code>
///   <para>Uses the OfficeIMO PowerPoint visual review profile.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePowerPointHtml", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("ConvertTo-PowerPointHtml")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficePowerPointHtmlCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetPresentation = "Presentation";

    /// <summary>Path to the presentation to convert.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Presentation instance to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetPresentation)]
    public PowerPointPresentation Presentation { get; set; } = null!;

    /// <summary>Password used to open an encrypted presentation package.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? Password { get; set; }

    /// <summary>Optional output HTML path. When omitted, HTML text is returned.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>HTML conversion profile.</summary>
    [Parameter]
    public OfficePowerPointHtmlProfile Profile { get; set; } = OfficePowerPointHtmlProfile.SemanticSlides;

    /// <summary>Built-in HTML document theme.</summary>
    [Parameter]
    public OfficeVisualThemeKind Theme { get; set; } = OfficeVisualThemeKind.WordLike;

    /// <summary>Optional HTML document title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Include hidden slides in the HTML review output.</summary>
    [Parameter]
    public SwitchParameter IncludeHiddenSlides { get; set; }

    /// <summary>Do not include presenter notes.</summary>
    [Parameter]
    public SwitchParameter NoNotes { get; set; }

    /// <summary>Do not include table content.</summary>
    [Parameter]
    public SwitchParameter NoTables { get; set; }

    /// <summary>Include hidden shapes.</summary>
    [Parameter]
    public SwitchParameter IncludeHiddenShapes { get; set; }

    /// <summary>Do not include extraction proof metadata.</summary>
    [Parameter]
    public SwitchParameter NoExtractionProof { get; set; }

    /// <summary>Do not include OfficeIMO default CSS styles.</summary>
    [Parameter]
    public SwitchParameter NoDefaultStyles { get; set; }

    /// <summary>Emit a FileInfo when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        PowerPointPresentation? presentation = null;
        var dispose = false;
        try
        {
            presentation = ResolvePresentation(out dispose);
            var html = presentation.ToHtml(CreateOptions());
            WriteHtml(html);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ConvertToOfficePowerPointHtmlFailed", ErrorCategory.InvalidOperation,
                ParameterSetName == ParameterSetPath ? Path : Presentation));
        }
        finally
        {
            if (dispose && presentation != null)
            {
                PowerPointDocumentService.ClosePresentation(presentation, save: false, show: false);
            }
        }
    }

    private PowerPointPresentation ResolvePresentation(out bool dispose)
    {
        dispose = false;
        if (ParameterSetName == ParameterSetPresentation)
        {
            return Presentation ?? throw new PSArgumentNullException(nameof(Presentation));
        }

        dispose = true;
        return PowerPointDocumentService.LoadPresentation(PdfCommandUtilities.ResolvePath(this, Path), Password, readOnly: true);
    }

    private PowerPointHtmlSaveOptions CreateOptions()
    {
        var options = new PowerPointHtmlSaveOptions
        {
            Profile = Profile == OfficePowerPointHtmlProfile.VisualReview
                ? OfficeHtmlConversionProfile.PowerPointVisualReview
                : OfficeHtmlConversionProfile.PowerPointSemanticSlides,
            Theme = Theme,
            IncludeDefaultStyles = !NoDefaultStyles.IsPresent,
            IncludeHiddenSlides = IncludeHiddenSlides.IsPresent,
            IncludeNotes = !NoNotes.IsPresent,
            IncludeTables = !NoTables.IsPresent,
            IncludeHiddenShapes = IncludeHiddenShapes.IsPresent,
            IncludeExtractionProof = !NoExtractionProof.IsPresent
        };

        if (!string.IsNullOrWhiteSpace(Title))
        {
            options.Title = Title!;
        }

        return options;
    }

    private void WriteHtml(string html)
    {
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteObject(html);
            return;
        }

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write PowerPoint HTML"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        File.WriteAllText(outputPath, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }
}

/// <summary>PowerPoint HTML conversion profiles exposed by PSWriteOffice.</summary>
public enum OfficePowerPointHtmlProfile
{
    /// <summary>Emit semantic slide content and inventories.</summary>
    SemanticSlides,

    /// <summary>Emit a visual review artifact.</summary>
    VisualReview
}
