using System.Linq;
using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Lists OfficeIMO.Reader capabilities registered in the current PSWriteOffice process.</summary>
/// <example>
///   <summary>Show registered Reader adapters.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$capabilities = Get-OfficeDocumentCapability
/// $capabilities | Sort-Object Id | Select-Object Id, Extensions</code>
///   <para>Lists built-in and modular Reader handlers, including adapters such as PDF, RTF, HTML, CSV, JSON, XML, YAML, ZIP, EPUB, and Visio when available.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentCapability")]
[OutputType(typeof(ReaderHandlerCapability), typeof(ReaderCapabilityManifest))]
public sealed class GetOfficeDocumentCapabilityCommand : OfficeDocumentReaderCommandBase
{
    /// <summary>Return the capability manifest envelope instead of individual handlers.</summary>
    [Parameter]
    public SwitchParameter Manifest { get; set; }

    /// <summary>Return handlers supplied by OfficeIMO or registered by a caller.</summary>
    [Parameter]
    public ReaderHandlerOrigin? Origin { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var capabilities = EffectiveReader.GetCapabilities()
            .Where(capability => !Origin.HasValue || capability.Origin == Origin.Value)
            .ToArray();

        if (Manifest.IsPresent)
        {
            var manifest = EffectiveReader.GetCapabilityManifest();
            manifest.Handlers = capabilities;
            WriteObject(manifest);
            return;
        }

        WriteObject(capabilities, enumerateCollection: true);
    }
}
