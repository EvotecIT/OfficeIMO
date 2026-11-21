using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Creates a live TOML document, or atomically publishes a completed content block.</summary>
/// <para>Inside the block, Set-OfficeTomlNode and the other node commands edit the new document without -Document.
/// Setting a key below a missing table adds the table's [header] section.</para>
/// <example><summary>Write a project file in one block.</summary>
/// <code>New-OfficeToml -Path ./pyproject.toml {
///     Set-OfficeTomlNode -Select 'project.name' -Value 'reporting'
///     Set-OfficeTomlNode -Select 'tool.ruff.line-length' -Value 100
/// }</code></example>
/// <example><summary>Compose a live document and save it later.</summary>
/// <code>$document = New-OfficeToml
/// Set-OfficeTomlNode -Document $document -Select 'title' -Value 'Release notes'
/// Save-OfficeToml -Document $document -Path ./config.toml
/// Close-OfficeToml -Document $document</code></example>
[Cmdlet(VerbsCommon.New, "OfficeToml", DefaultParameterSetName = "Live", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(TomlDocument))]
public sealed class NewOfficeTomlCommand : OfficeNewLifecycleCmdlet<TomlDocument> {
    /// <summary>Allow an existing destination; a failed block preserves its original contents.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        RunDocument(Path, isNew: true, Force.IsPresent,
            path => path == null ? TomlDocument.Create() : TomlDocument.Create(path),
            (document, path, _) => { document.Save(path); return null; });
    }
}
