using OfficeIMO.Html;

namespace OfficeIMO.OneNote.Html;

/// <summary>Controls ordinary HTML import into typed offline OneNote models.</summary>
public sealed class HtmlToOneNoteOptions {
    private HtmlImportLimits _limits = HtmlImportLimits.CreateDefault();

    /// <summary>Shared native-artifact limits for this import operation.</summary>
    public HtmlImportLimits Limits {
        get => _limits;
        set => _limits = value ?? HtmlImportLimits.CreateDefault();
    }

    /// <summary>Name assigned to the generated section.</summary>
    public string SectionName { get; set; } = "Imported";

    /// <summary>Name assigned to the generated notebook.</summary>
    public string NotebookName { get; set; } = "Imported";

    /// <summary>Whether bounded data URI images are restored as native OneNote images.</summary>
    public bool ImportImages { get; set; } = true;

    internal HtmlToOneNoteOptions Clone() => new HtmlToOneNoteOptions {
        Limits = Limits.Clone(),
        SectionName = SectionName,
        NotebookName = NotebookName,
        ImportImages = ImportImages
    };
}
