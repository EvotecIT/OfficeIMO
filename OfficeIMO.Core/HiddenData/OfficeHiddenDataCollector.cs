using OfficeIMO.Provenance;

namespace OfficeIMO.Core.Internal;

/// <summary>Collects hidden-data findings and applies the selected removals as the format walks its document.</summary>
internal sealed class OfficeHiddenDataCollector {
    /// <summary>The sentence shown for a reason when the format has nothing more specific to say.</summary>
    internal static string? NoteFor(OfficeHiddenDataReason reason) => reason switch {
        OfficeHiddenDataReason.NotSelected => "Kept because this category wasn't chosen for removal.",
        OfficeHiddenDataReason.NeededByDocument => "Kept because removing it would break or change the document.",
        OfficeHiddenDataReason.ProtectsContentCredentials => "Kept because removing this picture's metadata would break its Content Credentials. Remove provenance too to clean both.",
        OfficeHiddenDataReason.SourceTypeKept => "Removed, except the label that says how the picture was made. Remove provenance too to drop the label.",
        OfficeHiddenDataReason.ContentCredentialsNoLongerVerify => "Kept, but the picture changed, so these Content Credentials no longer verify.",
        OfficeHiddenDataReason.WholeFileCredentialsCannotSurvive => "Removed because Content Credentials for the whole file can't stay valid once anything in the file changes. Remove only provenance to leave the rest of the file untouched.",
        OfficeHiddenDataReason.UnsupportedImageFormat => "Kept because this picture format can't be cleaned without re-encoding the picture.",
        OfficeHiddenDataReason.KeptByProvenanceChoices => "Kept by your provenance choices, or because it is malformed or doesn't declare AI generation.",
        OfficeHiddenDataReason.UnusedImage => "This picture is stored in the file but not used by the document. Cleaning removes the whole picture.",
        _ => null
    };

    private readonly List<OfficeHiddenDataFinding> _findings = new();

    internal OfficeHiddenDataCollector(string format, OfficeHiddenDataOptions? removal) {
        Format = format;
        Removal = removal;
    }

    internal string Format { get; }

    /// <summary>Removal options, or null when only inspecting.</summary>
    internal OfficeHiddenDataOptions? Removal { get; }

    internal bool Removes(OfficeHiddenDataKind kind) => Removal?.Removes(kind) == true;

    internal OfficeTrackedChangesAction TrackedChanges => Removal?.TrackedChanges ?? OfficeTrackedChangesAction.Keep;

    /// <summary>Records a finding. When its category is selected, runs <paramref name="remove"/> and records <paramref name="removedOutcome"/>; a null action retains it with <paramref name="retainedNote"/>.</summary>
    internal void Add(OfficeHiddenDataKind kind, string location, string description, string? value = null,
        Action? remove = null, string? retainedNote = null, OfficeHiddenDataOutcome removedOutcome = OfficeHiddenDataOutcome.Removed) {
        if (Removal == null) {
            Record(kind, location, description, value, OfficeHiddenDataOutcome.Found, null);
        } else if (!Removal.Removes(kind)) {
            Record(kind, location, description, value, OfficeHiddenDataOutcome.Retained, null, reason: OfficeHiddenDataReason.NotSelected);
        } else if (remove == null) {
            Record(kind, location, description, value, OfficeHiddenDataOutcome.Retained, retainedNote, reason: OfficeHiddenDataReason.NeededByDocument);
        } else {
            remove();
            Record(kind, location, description, value, removedOutcome, null);
        }
    }

    /// <summary>Records a finding whose outcome the format decided itself.</summary>
    /// <remarks>A null <paramref name="note"/> takes the standard sentence for <paramref name="reason"/>; a retained finding without a reason is <see cref="OfficeHiddenDataReason.NeededByDocument"/>.</remarks>
    internal void Record(OfficeHiddenDataKind kind, string location, string description, string? value,
        OfficeHiddenDataOutcome outcome, string? note, OfficeProvenanceEvidence? provenance = null,
        OfficeHiddenDataScope scope = OfficeHiddenDataScope.Document, OfficeHiddenDataReason reason = OfficeHiddenDataReason.None,
        string? part = null) {
        if (reason == OfficeHiddenDataReason.None && outcome == OfficeHiddenDataOutcome.Retained) {
            reason = note == NotSelectedText || note == TrackedChangesKeptText ? OfficeHiddenDataReason.NotSelected : OfficeHiddenDataReason.NeededByDocument;
        }
        _findings.Add(new OfficeHiddenDataFinding(kind, location, description, value, outcome, note ?? NoteFor(reason), provenance, scope, reason, part));
    }

    /// <summary>Why tracked changes stay when the caller keeps them.</summary>
    internal const string TrackedChangesKeptText = "Kept because tracked changes are only accepted or rejected when you choose to.";

    /// <summary>The standard not-selected sentence, for formats that record retained findings themselves.</summary>
    internal static readonly string NotSelectedText = NoteFor(OfficeHiddenDataReason.NotSelected)!;

    /// <summary>The outcome for a finding the format removes itself: found when inspecting, retained when its category is not selected, otherwise <paramref name="removed"/>.</summary>
    internal OfficeHiddenDataOutcome Resolve(OfficeHiddenDataKind kind, bool removed, OfficeHiddenDataReason retainedReason, out OfficeHiddenDataReason reason) {
        reason = OfficeHiddenDataReason.None;
        if (Removal == null) return OfficeHiddenDataOutcome.Found;
        if (!Removal.Removes(kind)) {
            reason = OfficeHiddenDataReason.NotSelected;
            return OfficeHiddenDataOutcome.Retained;
        }
        if (removed) return OfficeHiddenDataOutcome.Removed;
        reason = retainedReason;
        return OfficeHiddenDataOutcome.Retained;
    }

    internal OfficeHiddenDataReport Build() => new(Format, _findings.ToList());
}
