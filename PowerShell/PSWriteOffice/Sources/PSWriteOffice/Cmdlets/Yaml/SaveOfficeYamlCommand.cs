using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Saves a live YAML document and keeps it available.</summary>
/// <para>The default save is minimal-change: untouched text is copied from the source, in the encoding and byte-order
/// mark it was loaded with, and only edited nodes are written anew. WriteOptions with Reformat rewrites the whole
/// document through the writer.</para>
/// <example><summary>Save to the file the document was opened from.</summary>
/// <code>Save-OfficeYaml -Document $document</code></example>
/// <example><summary>Write a reformatted copy with four-space mappings.</summary>
/// <code>Save-OfficeYaml -Document $document -Path ./out/values.yaml -WriteOptions @{ Reformat = $true; WriterOptions = [OfficeIMO.Yaml.YamlWriterOptions]@{ MappingIndent = 4 } } -PassThru</code></example>
[Cmdlet(VerbsData.Save, "OfficeYaml", SupportsShouldProcess = true, DefaultParameterSetName = "Document")]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficeYamlCommand : OfficeCmdlet {
    /// <summary>YAML document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public YamlDocument Document { get; set; } = null!;
    /// <summary>Optional destination; omission saves to the file the document was opened or created with.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }
    /// <summary>Native save settings or a hashtable of their properties.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(YamlSaveOptions)), ValidateNotNull]
    public YamlSaveOptions? WriteOptions { get; set; }
    /// <summary>Return the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string? destination = Path ?? OfficeYamlDocuments.GetPath(Document);
        if (string.IsNullOrWhiteSpace(destination)) {
            InvalidArgument(new PSArgumentException("Supply Path when the document has no associated file."), Document);
            return;
        }
        destination = ResolveOutputPath(destination!);
        if (!ShouldProcess(destination, "Save YAML document")) return;
        try {
            OfficeOutput.Write(destination, stream => Document.Save(stream, WriteOptions), noClobber: false);
            // A document created without a path takes the first file it is saved to, as New- promises.
            if (OfficeYamlDocuments.GetPath(Document) == null) OfficeYamlDocuments.Associate(Document, destination);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) { WriteError(OfficeYamlDocuments.Error(error, destination)); }
    }
}
