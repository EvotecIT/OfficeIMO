using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Gets nodes of a live TOML document by path, as Path, Value, Node, Line and Column objects.</summary>
/// <para>Select uses the engine's path grammar, the one YAML node paths use: dotted keys, [n] for 0-based array items
/// and tables of an array of tables, and ["key.with.dots"] for keys that are not plain names. Value holds plain
/// PowerShell data as ConvertFrom-OfficeToml returns it; Node is the live node. Line and Column are 1-based and point at
/// the value or the [ of a header; tables made by dotted keys and nodes added by editing report 0. Without Select, the
/// root table is returned.</para>
/// <example><summary>Read one value.</summary>
/// <code>(Get-OfficeTomlNode -Document $document -Select 'tool.ruff.line-length').Value</code></example>
/// <example><summary>Find where settings live in the file.</summary>
/// <code>Get-OfficeTomlNode -Document $document -Select 'project.version', 'tool.uv.sources["my.pkg"]' |
///     Select-Object -Property Path, Value, Line, Column</code></example>
[Cmdlet(VerbsCommon.Get, "OfficeTomlNode", DefaultParameterSetName = "Document")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeTomlNodeCommand : OfficeDocumentCmdlet<TomlDocument> {
    /// <summary>Paths of the nodes to get; omission returns the root table.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string[]? Select { get; set; }
    /// <summary>Return tables in Value as ordered hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TomlDocument document = GetDocument();
        if (Select == null) {
            WriteObject(OfficeTomlDocuments.DescribeNode(document.Root, "$", AsHashtable.IsPresent));
            return;
        }
        foreach (string path in Select) {
            try {
                WriteObject(OfficeTomlDocuments.DescribeNode(OfficeTomlDocuments.Select(document, path), path, AsHashtable.IsPresent));
            } catch (PipelineStoppedException) { throw; }
            catch (Exception error) { WriteError(OfficeTomlDocuments.Error(error, path)); }
        }
    }
}
