using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Saves a live TOML document and keeps it available.</summary>
/// <para>The save is minimal-change: untouched text is copied from the source, with its line endings and UTF-8
/// byte-order mark, and only edited keys, values and comments are written anew in the layout of their neighbors.</para>
/// <example><summary>Save to the file the document was opened from.</summary>
/// <code>Save-OfficeToml -Document $document</code></example>
/// <example><summary>Save to a new file, which the document then saves to, and return it.</summary>
/// <code>Save-OfficeToml -Document $document -Path ./out/pyproject.toml -PassThru</code></example>
[Cmdlet(VerbsData.Save, "OfficeToml", SupportsShouldProcess = true, DefaultParameterSetName = "Document")]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficeTomlCommand : OfficeCmdlet {
    /// <summary>TOML document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public TomlDocument Document { get; set; } = null!;
    /// <summary>Optional destination, which becomes the document's FilePath; omission saves to its FilePath.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }
    /// <summary>Return the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string? destination = Path ?? Document.FilePath;
        if (string.IsNullOrWhiteSpace(destination)) {
            InvalidArgument(new PSArgumentException("Supply Path when the document has no associated file."), Document);
            return;
        }
        destination = ResolveOutputPath(destination!);
        if (!ShouldProcess(destination, "Save TOML document")) return;
        try {
            Document.Save(destination);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) { WriteError(OfficeTomlDocuments.Error(error, destination)); }
    }
}
