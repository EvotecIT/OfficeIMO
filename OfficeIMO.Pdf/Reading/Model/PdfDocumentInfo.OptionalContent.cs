namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentInfo {
    private IReadOnlyList<string>? _optionalContentGroupNames;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfOptionalContentGroup>>? _optionalContentGroupsByName;

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
    public IReadOnlyList<string> OptionalContentGroupNames {
        get {
            if (_optionalContentGroupNames is not null) {
                return _optionalContentGroupNames;
            }

            var names = new List<string>(OptionalContentGroups.Count);
            for (int i = 0; i < OptionalContentGroups.Count; i++) {
                names.Add(OptionalContentGroups[i].Name);
            }

            _optionalContentGroupNames = names.AsReadOnly();
            return _optionalContentGroupNames;
        }
    }

    /// <summary>Optional-content groups grouped by layer display name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfOptionalContentGroup>> OptionalContentGroupsByName {
        get {
            if (_optionalContentGroupsByName is not null) {
                return _optionalContentGroupsByName;
            }

            var grouped = new Dictionary<string, List<PdfOptionalContentGroup>>(StringComparer.Ordinal);
            for (int i = 0; i < OptionalContentGroups.Count; i++) {
                PdfOptionalContentGroup group = OptionalContentGroups[i];
                if (!grouped.TryGetValue(group.Name, out List<PdfOptionalContentGroup>? groups)) {
                    groups = new List<PdfOptionalContentGroup>();
                    grouped.Add(group.Name, groups);
                }

                groups.Add(group);
            }

            _optionalContentGroupsByName = ToReadOnlyLookup(grouped);
            return _optionalContentGroupsByName;
        }
    }

    /// <summary>Returns optional-content groups with a matching layer display name.</summary>
    public IReadOnlyList<PdfOptionalContentGroup> GetOptionalContentGroupsByName(string name) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        return OptionalContentGroupsByName.TryGetValue(name, out IReadOnlyList<PdfOptionalContentGroup>? groups)
            ? groups
            : Array.Empty<PdfOptionalContentGroup>();
    }
}
