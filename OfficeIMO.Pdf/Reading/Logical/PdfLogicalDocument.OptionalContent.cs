namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalDocument {
    /// <summary>Catalog optional-content/layer metadata discovered from /OCProperties.</summary>
    public PdfOptionalContentProperties? OptionalContent { get; }

    /// <summary>True when readable optional-content/layer metadata was discovered.</summary>
    public bool HasReadableOptionalContent => OptionalContent is not null;

    /// <summary>Optional-content groups discovered from catalog /OCProperties.</summary>
    public IReadOnlyList<PdfOptionalContentGroup> OptionalContentGroups => OptionalContent?.Groups ?? Array.Empty<PdfOptionalContentGroup>();

    /// <summary>Number of optional-content groups discovered from catalog /OCProperties.</summary>
    public int OptionalContentGroupCount => OptionalContentGroups.Count;

    /// <summary>True when at least one optional-content group was discovered.</summary>
    public bool HasOptionalContentGroups => OptionalContentGroupCount > 0;

    /// <summary>Optional-content group names in first-seen catalog order.</summary>
    public IReadOnlyList<string> OptionalContentGroupNames => OptionalContentGroups.Select(group => group.Name).ToArray();

    /// <summary>Returns optional-content groups with a matching layer display name.</summary>
    public IReadOnlyList<PdfOptionalContentGroup> GetOptionalContentGroupsByName(string name) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        return OptionalContentGroups.Where(group => string.Equals(group.Name, name, StringComparison.Ordinal)).ToArray();
    }
}
