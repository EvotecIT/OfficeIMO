using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Generates the OfficeIMO Visio reference gallery as editable .vsdx diagrams.</summary>
/// <example>
///   <summary>Generate the Visio reference gallery.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeVisioGallery -OutputDirectory .\VisioGallery |
///     Select-Object Name, FilePath, IsClean</code>
///   <para>Creates polished, editable Visio samples for flowcharts, architecture, network, timeline, swimlane, org chart, and graph diagrams.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeVisioGallery")]
[Alias("New-VisioGallery")]
[OutputType(typeof(VisioGalleryResult))]
public sealed class NewOfficeVisioGalleryCommand : PSCmdlet
{
    /// <summary>Directory that receives generated .vsdx gallery documents.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>Skip structural package validation after gallery documents are generated.</summary>
    [Parameter]
    public SwitchParameter NoPackageValidation { get; set; }

    /// <summary>Skip visual quality analysis after gallery documents are generated.</summary>
    [Parameter]
    public SwitchParameter NoVisualQualityAnalysis { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var outputDirectory = VisioCommandUtilities.ResolvePath(this, OutputDirectory);
        var results = VisioGallery.Create(outputDirectory, new VisioGalleryOptions
        {
            ValidatePackage = !NoPackageValidation.IsPresent,
            AnalyzeVisualQuality = !NoVisualQualityAnalysis.IsPresent
        });

        WriteObject(results, enumerateCollection: true);
    }
}
