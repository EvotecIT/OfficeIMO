using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Runs package, preflight, accessibility, feature, review, animation, signature, and visual inspections.</summary>
/// <example>
///   <summary>Run accessibility and layout preflight.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointInspection -Path .\Deck.pptx</code>
///   <para>Returns one coherent inspection report over the same presentation model used for editing.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointInspection", DefaultParameterSetName = "Path")]
[OutputType(typeof(PowerPointInspectionReport))]
public sealed class GetOfficePowerPointInspectionCommand : PSCmdlet
{
    /// <summary>Path to the presentation.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open presentation instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Presentation")]
    public PowerPointPresentation Presentation { get; set; } = null!;

    /// <summary>Optional report selection and inspection policies.</summary>
    [Parameter]
    public PowerPointInspectionOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        PowerPointPresentation? owned = null;
        try
        {
            var presentation = Presentation;
            if (ParameterSetName == "Path")
            {
                owned = PowerPointDocumentService.LoadPresentation(
                    SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path),
                    readOnly: true);
                presentation = owned;
            }
            WriteObject(presentation.Inspect(Options));
        }
        finally
        {
            if (owned != null)
            {
                PowerPointDocumentService.ClosePresentation(owned, save: false, show: false);
            }
        }
    }
}
