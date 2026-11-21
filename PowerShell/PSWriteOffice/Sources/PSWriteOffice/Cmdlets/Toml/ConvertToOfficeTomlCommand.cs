using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Converts PowerShell objects to TOML text, as ConvertTo-Json does for JSON.</summary>
/// <para>Hashtables, ordered dictionaries and object properties become tables: nested objects are written as
/// <c>[table]</c> sections and lists of objects as <c>[[array]]</c> sections, after the table's own values. Collections
/// become arrays, DateTime and DateTimeOffset TOML date-times, and enums their names. TOML has no null, so nulls are
/// omitted with a warning, and input that is not an object (such as several piped objects) is written under RootName.
/// Objects nested deeper than Depth (default 2) are written as their string form with a warning.</para>
/// <example><summary>Turn a hashtable into a configuration file's text.</summary>
/// <code>[ordered]@{ name = 'web'; port = 8080; tls = [ordered]@{ enabled = $true } } | ConvertTo-OfficeToml</code></example>
/// <example><summary>Write a list of records as an array of tables.</summary>
/// <code>Get-Process | Select-Object -First 3 Name, Id | ConvertTo-OfficeToml -RootName process</code></example>
[Cmdlet(VerbsData.ConvertTo, "OfficeToml", DefaultParameterSetName = "InputObject")]
[OutputType(typeof(string))]
public sealed class ConvertToOfficeTomlCommand : OfficeCmdlet {
    private readonly OfficeTomlWriterInput _input = new();

    /// <summary>Objects to convert; an argument is written exactly as given, piped objects as one value or an array.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "InputObject"), AllowNull, AllowEmptyCollection]
    public object? InputObject { get; set; }
    /// <summary>How many levels of nested objects to write; deeper values become their string form.</summary>
    [Parameter, ValidateRange(0, TomlPowerShellValue.MaxDepth)]
    public int Depth { get; set; } = 2;
    /// <summary>The key that holds input which is not an object, because a TOML document is a table.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string RootName { get; set; } = "root";

    /// <inheritdoc />
    protected override void ProcessRecord() => _input.Add(this, InputObject);

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (!_input.HasInput) return;
        try {
            WriteObject(OfficeTomlDocuments.Serialize(this, _input.Value, Depth, RootName));
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) when (error is TomlException or ArgumentException or InvalidOperationException or NotSupportedException) {
            ThrowTerminatingError(OfficeTomlDocuments.Error(error, InputObject));
        }
    }
}
