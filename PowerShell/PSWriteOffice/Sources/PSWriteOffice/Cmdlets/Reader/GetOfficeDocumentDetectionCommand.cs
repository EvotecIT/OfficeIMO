using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Detects a document kind from extension and bounded content evidence.</summary>
/// <example>
///   <summary>Inspect the real content type of a document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeDocumentDetection -Path .\upload.bin -Mode PreferContent</code>
///   <para>Returns the selected kind, confidence, media type, and evidence used by OfficeIMO.Reader.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentDetection")]
[OutputType(typeof(ReaderDetectionResult))]
public sealed class GetOfficeDocumentDetectionCommand : OfficeDocumentReaderCommandBase
{
    /// <summary>Path to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Policy used to combine extension and content evidence.</summary>
    [Parameter]
    public ReaderDetectionMode Mode { get; set; } = ReaderDetectionMode.PreferContent;

    /// <summary>Maximum prefix bytes inspected for signatures and text markers.</summary>
    [Parameter]
    public int? MaxProbeBytes { get; set; }

    /// <summary>Maximum archive entries inspected while classifying container formats.</summary>
    [Parameter]
    public int? MaxContainerEntries { get; set; }

    /// <summary>Skip structural inspection of ZIP-based containers.</summary>
    [Parameter]
    public SwitchParameter NoContainerInspection { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var options = new ReaderDetectionOptions
        {
            Mode = Mode,
            InspectContainers = !NoContainerInspection.IsPresent
        };
        if (MaxProbeBytes.HasValue) options.MaxProbeBytes = MaxProbeBytes.Value;
        if (MaxContainerEntries.HasValue) options.MaxContainerEntries = MaxContainerEntries.Value;
        WriteObject(EffectiveReader.Detect(ReaderCommandUtilities.ResolvePath(this, Path), options));
    }
}
