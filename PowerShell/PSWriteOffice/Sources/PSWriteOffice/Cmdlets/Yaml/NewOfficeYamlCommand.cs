using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Creates a live YAML document, or atomically publishes a completed content block.</summary>
/// <para>Inside the block, Set-OfficeYamlNode and the other node commands edit the new document without -Document.</para>
/// <example><summary>Write a configuration file in one block.</summary>
/// <code>New-OfficeYaml -Path ./service.yaml {
///     Set-OfficeYamlNode -Select 'name' -Value 'web'
///     Set-OfficeYamlNode -Select 'spec.replicas' -Value 3
/// }</code></example>
/// <example><summary>Compose a live document and save it later.</summary>
/// <code>$document = New-OfficeYaml
/// Set-OfficeYamlNode -Document $document -Select 'image' -Value 'nginx:1.27'
/// Save-OfficeYaml -Document $document -Path ./values.yaml
/// Close-OfficeYaml -Document $document</code></example>
[Cmdlet(VerbsCommon.New, "OfficeYaml", DefaultParameterSetName = "Live", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(YamlDocument))]
public sealed class NewOfficeYamlCommand : OfficeNewLifecycleCmdlet<YamlDocument> {
    /// <summary>Allow an existing destination; a failed block preserves its original contents.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }
    /// <summary>Native YAML save settings or a hashtable, used when publishing a block.</summary>
    [Parameter(ParameterSetName = "Content"), OfficeOptionsTransformation(typeof(YamlSaveOptions)), ValidateNotNull]
    public YamlSaveOptions? WriteOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        RunDocument(Path, isNew: true, Force.IsPresent,
            path => path == null ? YamlDocument.Create() : YamlDocument.Create(path),
            (document, path, _) => { document.Save(path, WriteOptions); return null; });
    }
}
