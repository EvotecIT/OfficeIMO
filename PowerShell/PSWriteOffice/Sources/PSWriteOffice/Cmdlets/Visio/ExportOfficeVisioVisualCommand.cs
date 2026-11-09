using System.Management.Automation;
using OfficeIMO.ChartForgeX;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Visio;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Exports CFX semantic visual-artifact input as a native editable VSDX diagram.</summary>
/// <example>
///   <summary>Export an ImagePlayground topology directly to VSDX.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$topology | ConvertTo-ImageVisualArtifact | Export-OfficeVisioVisual -Path .\Topology.vsdx</code>
///   <para>Creates native Visio shapes, containers, connectors, Shape Data, and hyperlinks from the portable CFX semantics.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeVisioVisual", SupportsShouldProcess = true, DefaultParameterSetName = "InputObject")]
[OutputType(typeof(OfficeVisioVisualConversionResult), typeof(FileInfo))]
public sealed class ExportOfficeVisioVisualCommand : OfficeVisioVisualCommandBase {
    private object? _bufferedInput;
    private bool _inputSeen;

    /// <summary>Typed CFX artifact, semantic envelope/JSON, ImagePlayground portable artifact, or prior conversion result.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, ParameterSetName = "InputObject")]
    public object InputObject { get; set; } = null!;

    /// <summary>Destination .vsdx path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Show the generated VSDX after saving.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Refuse to replace an existing VSDX destination.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Emit the conversion result instead of the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (BufferPipelineByte(InputObject)) {
            return;
        }

        if (_inputSeen) {
            throw new PSInvalidOperationException(
                "Export-OfficeVisioVisual accepts one input artifact for one output path. Invoke the cmdlet separately for each destination.");
        }
        _inputSeen = true;
        _bufferedInput = InputObject;
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        byte[]? jsonBytes = CompletePipelineBytes();
        if (jsonBytes != null) {
            _bufferedInput = jsonBytes;
            _inputSeen = true;
        }

        if (!_inputSeen) {
            return;
        }

        string fullPath = VisioCommandUtilities.ResolvePath(this, Path);
        if (!string.Equals(System.IO.Path.GetExtension(fullPath), ".vsdx", System.StringComparison.OrdinalIgnoreCase)) {
            throw new PSArgumentException("Native editable Visio output must use the .vsdx extension.", nameof(Path));
        }
        if (!ShouldProcess(fullPath, "Export CFX visual artifact as native editable Visio")) {
            return;
        }

        OfficePath.ValidateDestination(fullPath, NoClobber);
        object? supplied = _bufferedInput is PSObject wrapped ? wrapped.BaseObject : _bufferedInput;
        bool ownsDocument = supplied is not OfficeVisioVisualConversionResult;
        OfficeVisioVisualConversionResult result = ResolveVisioVisual(_bufferedInput!);
        VisioCommandUtilities.EnsureDirectory(fullPath);
        try {
            OfficeOutput.Write(fullPath, stream => result.Document.Save(stream), NoClobber);
        } finally {
            if (ownsDocument && !PassThru) result.Document.Dispose();
        }
        if (Show.IsPresent) {
            FileOpenService.Open(fullPath);
        }
        WriteObject(PassThru.IsPresent ? result : new FileInfo(fullPath));
    }
}