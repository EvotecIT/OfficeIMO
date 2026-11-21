using System.Management.Automation;
using System.Text;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Writes PowerShell objects to a TOML file; the file form of ConvertTo-OfficeToml.</summary>
/// <para>The file is UTF-8 without a byte-order mark and is written atomically. An existing file is replaced unless
/// NoClobber is set. Nested objects become <c>[table]</c> sections and lists of objects <c>[[array]]</c> sections; nulls
/// are omitted with a warning.</para>
/// <example><summary>Save settings to a file.</summary>
/// <code>[ordered]@{ title = 'Report'; owner = [ordered]@{ name = 'Ops'; updated = Get-Date } } | Export-OfficeToml -Path ./settings.toml</code></example>
/// <example><summary>Export an inventory without replacing an existing file.</summary>
/// <code>Import-OfficeCsv -Path ./servers.csv | Export-OfficeToml -Path ./inventory.toml -RootName server -NoClobber -PassThru</code></example>
[Cmdlet(VerbsData.Export, "OfficeToml", DefaultParameterSetName = "InputObject", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeTomlCommand : OfficeCmdlet {
    private readonly OfficeTomlWriterInput _input = new();

    /// <summary>Objects to write; an argument is written exactly as given, piped objects as one value or an array.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "InputObject"), AllowNull, AllowEmptyCollection]
    public object? InputObject { get; set; }
    /// <summary>The TOML file to create.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;
    /// <summary>How many levels of nested objects to write; deeper values become their string form.</summary>
    [Parameter, ValidateRange(0, TomlPowerShellValue.MaxDepth)]
    public int Depth { get; set; } = 2;
    /// <summary>The key that holds input which is not an object, because a TOML document is a table.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string RootName { get; set; } = "root";
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
        if (!ShouldProcess(destination, "Export TOML")) return;
        try {
            byte[] bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(OfficeTomlDocuments.Serialize(this, _input.Value, Depth, RootName));
            OfficeOutput.Write(destination, stream => stream.Write(bytes, 0, bytes.Length), NoClobber.IsPresent);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) { ThrowTerminatingError(OfficeTomlDocuments.Error(error, destination)); }
    }
}
