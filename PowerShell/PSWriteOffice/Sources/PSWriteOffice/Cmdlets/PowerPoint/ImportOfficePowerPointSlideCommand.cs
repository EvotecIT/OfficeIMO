using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Imports a slide from another PowerPoint presentation.</summary>
/// <para>Can import from an open presentation or directly from a source file path.</para>
/// <example>
///   <summary>Import the first slide from another deck.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointImportTarget.pptx {
///     Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Target deck'
///     Import-OfficePowerPointSlide -SourcePath .\Examples\Documents\SourceDeck.pptx -SourceIndex 0 -InsertAt 1
/// }</code>
///   <para>Imports the first slide from another deck into the target presentation.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficePowerPointSlide", DefaultParameterSetName = ParameterSetSourcePresentation)]
[OutputType(typeof(PowerPointSlide))]
public sealed class ImportOfficePowerPointSlideCommand : PSCmdlet
{
    private const string ParameterSetSourcePresentation = "SourcePresentation";
    private const string ParameterSetSourcePath = "SourcePath";

    /// <summary>Target presentation to update (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Source presentation to import from.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSourcePresentation)]
    public PowerPointPresentation SourcePresentation { get; set; } = null!;

    /// <summary>Path to the source presentation.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSourcePath)]
    [Alias("Path")]
    public string SourcePath { get; set; } = string.Empty;

    /// <summary>Zero-based slide index in the source presentation.</summary>
    [Parameter(Mandatory = true)]
    public int SourceIndex { get; set; }

    /// <summary>Optional target insertion index; omit to append.</summary>
    [Parameter]
    public int? InsertAt { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var targetPresentation = Presentation ?? PowerPointDslContext.Current?.Presentation
                ?? throw new InvalidOperationException("Target presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

            if (ParameterSetName == ParameterSetSourcePresentation)
            {
                WriteObject(targetPresentation.ImportSlide(SourcePresentation, SourceIndex, InsertAt));
                return;
            }

            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(SourcePath);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            using var sourcePresentation = PowerPointPresentation.Load(resolvedPath);
            WriteObject(targetPresentation.ImportSlide(sourcePresentation, SourceIndex, InsertAt));
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointImportSlideFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
