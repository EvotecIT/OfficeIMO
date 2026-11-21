using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Yaml;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Tests whether YAML files or text load, without changing them.</summary>
/// <para>Returns $true or $false for each input, as Test-Json does. Detailed returns a result with the error's
/// diagnostic code, message, 1-based line and column, and the warnings the engine recorded while loading. Strict applies the
/// specification's exact rules, and Untrusted the engine's limits for files from outside.</para>
/// <example><summary>Gate a pipeline step on a valid file.</summary>
/// <code>if (-not (Test-OfficeYaml -Path ./values.yaml -Strict)) { throw 'values.yaml is not valid YAML.' }</code></example>
/// <example><summary>Report where every broken file fails.</summary>
/// <code>Get-ChildItem ./deploy -Recurse -Filter *.yaml | Test-OfficeYaml -Detailed |
///     Where-Object { -not $_.IsValid } |
///     Select-Object -Property Path, Line, Column, Code, Message</code></example>
[Cmdlet(VerbsDiagnostic.Test, "OfficeYaml", DefaultParameterSetName = "Path")]
[OutputType(typeof(bool), typeof(OfficeYamlTestResult))]
public sealed class TestOfficeYamlCommand : OfficeCmdlet {
    /// <summary>YAML files with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal YAML files, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>YAML text to test.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Native load options or a hashtable of their properties.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(YamlLoadOptions)), ValidateNotNull]
    public YamlLoadOptions? Options { get; set; }
    /// <summary>Caps every limit at the engine's untrusted defaults.</summary>
    [Parameter]
    public SwitchParameter Untrusted { get; set; }
    /// <summary>Applies the YAML specification's exact rules instead of accepting common real-world leniencies.</summary>
    [Parameter]
    public SwitchParameter Strict { get; set; }
    /// <summary>Forces a schema: Failsafe, Json, Core (the default) or Yaml11.</summary>
    [Parameter]
    public YamlSchema? Schema { get; set; }
    /// <summary>Return a result with the error position and the load warnings instead of a Boolean.</summary>
    [Parameter]
    public SwitchParameter Detailed { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        YamlLoadOptions options = OfficeYamlDocuments.BuildLoadOptions(Options, Untrusted.IsPresent, BoundSwitch(nameof(Strict)), Schema);
        if (ParameterSetName == "Text") {
            Write(null, () => YamlDocument.Parse(Text, options));
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string source in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) {
            Write(source, () => YamlDocument.Load(source, options));
        }
    }

    private void Write(string? path, Func<YamlDocument> load) {
        OfficeYamlTestResult result;
        try {
            using YamlDocument document = load();
            result = new OfficeYamlTestResult(path, document.Diagnostics);
        } catch (YamlException error) {
            result = new OfficeYamlTestResult(path, Array.Empty<OfficeDiagnostic>(), error);
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            // A file that cannot be read is not a verdict on its YAML; report it and test the rest.
            WriteError(OfficeYamlDocuments.Error(error, path));
            return;
        }
        if (Detailed) WriteObject(result);
        else WriteObject(result.IsValid);
    }
}
