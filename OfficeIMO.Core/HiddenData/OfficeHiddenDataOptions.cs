namespace OfficeIMO;

/// <summary>Selects which hidden-data categories <c>RemoveHiddenData</c> removes from a Word, Excel or PowerPoint document.</summary>
public sealed class OfficeHiddenDataOptions {
    private OfficeHiddenDataKind _kinds = OfficeHiddenDataKind.Default;
    private string _anonymousAuthor = "Author";

    /// <summary>Categories to remove. Defaults to <see cref="OfficeHiddenDataKind.Default"/>; tracked changes, hidden content and provenance are opt-in.</summary>
    public OfficeHiddenDataKind Kinds {
        get => _kinds;
        set {
            if ((value & ~OfficeHiddenDataKind.All) != 0) throw new ArgumentOutOfRangeException(nameof(Kinds), value, "Unknown hidden-data category.");
            _kinds = value;
        }
    }

    /// <summary>Whether tracked changes are kept, accepted or rejected. Defaults to <see cref="OfficeTrackedChangesAction.Keep"/>.</summary>
    public OfficeTrackedChangesAction TrackedChanges { get; set; } = OfficeTrackedChangesAction.Keep;

    /// <summary>Name that replaces authors on retained comments and revisions when <see cref="OfficeHiddenDataKind.PersonalIdentities"/> is selected. Defaults to <c>Author</c>.</summary>
    public string AnonymousAuthor {
        get => _anonymousAuthor;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("The anonymous author name must not be empty; Office requires an author name on comments and revisions.", nameof(value));
            _anonymousAuthor = value.Trim();
        }
    }

    internal bool Removes(OfficeHiddenDataKind kind) => kind != OfficeHiddenDataKind.TrackedChanges && (Kinds & kind) == kind;

    internal string AnonymousInitials => AnonymousAuthor.Substring(0, 1).ToUpperInvariant();
}
