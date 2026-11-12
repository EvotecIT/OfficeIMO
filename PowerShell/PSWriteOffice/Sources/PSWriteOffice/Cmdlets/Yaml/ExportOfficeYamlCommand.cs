using System.Management.Automation;
using System.Text;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Writes PowerShell objects to a YAML file; the file form of ConvertTo-OfficeYaml.</summary>
/// <para>The file is UTF-8 without a byte-order mark and is written atomically. An existing file is replaced unless
/// NoClobber is set. To change values in an existing file while keeping its comments and layout, use Open-OfficeYaml
/// and Set-OfficeYamlNode instead.</para>
/// <example><summary>Save settings to a file.</summary>
/// <code>[ordered]@{ logging = [ordered]@{ level = 'Information' }; retries = 3 } | Export-OfficeYaml -Path ./settings.yaml</code></example>
/// <example><summary>Export an inventory without replacing an existing file.</summary>
/// <code>Import-OfficeCsv -Path ./servers.csv | Export-OfficeYaml -Path ./inventory.yaml -NoClobber -PassThru</code></example>
[Cmdlet(VerbsData.Export, "OfficeYaml", DefaultParameterSetName = "InputObject", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeYamlCommand : OfficeCmdlet {
    private readonly OfficeYamlWriterInput _input = new();

    /// <summary>Objects to write; an argument is written exactly as given, piped objects as one value or a sequence.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "InputObject"), AllowNull, AllowEmptyCollection]
    public object? InputObject { get; set; }
    /// <summary>The YAML file to create.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;
    /// <summary>How many levels of nested objects to write; deeper values become their string form.</summary>
    [Parameter, ValidateRange(0, YamlPowerShellValue.MaxDepth)]
    public int Depth { get; set; } = 2;
    /// <summary>Indentation and line-break settings, as a YamlWriterOptions object or hashtable.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(YamlWriterOptions)), ValidateNotNull]
    public YamlWriterOptions? WriteOptions { get; set; }
    /// <summary>Refuse to replace an existing file.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }
    /// <summary>Return the written file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => _input.Add(this, InputObject);

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (!_input.HasInput) return;
        string destination = ResolveOutputPath(Path);
        if (!ShouldProcess(destination, "Export YAML")) return;
        try {
            byte[] bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(_input.Serialize(this, Depth, WriteOptions));
            OfficeOutput.Write(destination, stream => stream.Write(bytes, 0, bytes.Length), NoClobber.IsPresent);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) { ThrowTerminatingError(OfficeYamlDocuments.Error(error, destination)); }
    }
}
