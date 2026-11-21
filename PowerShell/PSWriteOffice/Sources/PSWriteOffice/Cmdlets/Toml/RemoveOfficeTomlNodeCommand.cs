using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Removes keys, tables or array items from a live TOML document by path.</summary>
/// <para>A key goes with its line and the comment lines directly above it; a table goes with all of its sections,
/// wherever they are in the file. Neighboring keys and comments stay. A path with nothing at it writes a
/// non-terminating ObjectNotFound error.</para>
/// <example><summary>Drop a setting inside an editing block.</summary>
/// <code>Open-OfficeToml -Path ./pyproject.toml {
///     Remove-OfficeTomlNode -Select 'tool.ruff.lint.ignore'
/// }</code></example>
/// <example><summary>Remove a table of an array of tables if it is there.</summary>
/// <code>Remove-OfficeTomlNode -Document $document -Select 'bin[1]' -ErrorAction SilentlyContinue</code></example>
[Cmdlet(VerbsCommon.Remove, "OfficeTomlNode", DefaultParameterSetName = "Document")]
public sealed class RemoveOfficeTomlNodeCommand : OfficeDocumentCmdlet<TomlDocument> {
    /// <summary>Paths of the keys, tables or items to remove.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNullOrEmpty]
    public string[] Select { get; set; } = Array.Empty<string>();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TomlDocument document = GetDocument();
        foreach (string path in Select) {
            try {
                if (!document.Remove(path)) throw new ItemNotFoundException("No TOML node is at '" + path + "'.");
            } catch (PipelineStoppedException) { throw; }
            catch (Exception error) { WriteError(OfficeTomlDocuments.Error(error, path)); }
        }
    }
}
