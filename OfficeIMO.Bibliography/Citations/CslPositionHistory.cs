namespace OfficeIMO.Bibliography;

/// <summary>Operation-local citation history, with separate body and note sequences.</summary>
internal sealed class CslPositionHistory {
    private readonly Sequence _body = new Sequence();
    private readonly Sequence _notes = new Sequence();
    private readonly HashSet<int> _renderedNotes = new HashSet<int>();
    private readonly int _nearDistance;
    private readonly bool _noteStyle;

    internal CslPositionHistory(bool noteStyle, int nearDistance) {
        _noteStyle = noteStyle;
        _nearDistance = nearDistance;
    }

    internal CslCitationItem? Preceding(CslCitation citation) {
        Sequence sequence = citation.NoteIndex > 0 ? _notes : _body;
        CslCitation? previous = sequence.Previous;
        if (previous == null || previous.Items.Count != 1) return null;
        if (citation.NoteIndex > 0 && previous.NoteIndex != citation.NoteIndex &&
            (citation.NoteIndex - previous.NoteIndex != 1 || sequence.NoteKeys.Count != 1)) return null;
        return previous.Items[0];
    }

    internal void SetPosition(CslContext context, CslCitation citation, CslCitationItem? preceding) {
        Sequence sequence = citation.NoteIndex > 0 ? _notes : _body;
        CslCitationItem current = context.Cite!;
        context.Position = sequence.Seen.Add(current.Key) ? "first" : "subsequent";
        context.NoteIndex = citation.NoteIndex;
        if (_noteStyle && citation.NoteIndex > 0 && sequence.LastNotes.TryGetValue(current.Key, out int lastNote)) {
            int distance = citation.NoteIndex - lastNote;
            context.NearNote = distance >= 0 && distance <= _nearDistance;
        }
        if (preceding != null && preceding.Key == current.Key) {
            bool hadLocator = !string.IsNullOrEmpty(preceding.Locator);
            bool hasLocator = !string.IsNullOrEmpty(current.Locator);
            if (!hadLocator && !hasLocator || hadLocator && hasLocator &&
                string.Equals(preceding.Locator, current.Locator, StringComparison.Ordinal) && preceding.LocatorType == current.LocatorType)
                context.Position = "ibid";
            else if (hasLocator) context.Position = "ibid-with-locator";
        }
        if (citation.NoteIndex > 0) sequence.LastNotes[current.Key] = citation.NoteIndex;
    }

    internal void Complete(CslCitation citation) {
        if (citation.Items.Count == 0) return;
        Sequence sequence = citation.NoteIndex > 0 ? _notes : _body;
        if (sequence.Previous?.NoteIndex != citation.NoteIndex) sequence.NoteKeys.Clear();
        foreach (CslCitationItem item in citation.Items) sequence.NoteKeys.Add(item.Key);
        sequence.Previous = citation;
    }

    internal bool StartsNote(CslCitation citation, CslCitationItem? firstVisible, CslText rendered) {
        if (!_noteStyle || citation.NoteIndex == 0 || rendered.IsEmpty) return false;
        bool first = _renderedNotes.Add(citation.NoteIndex);
        return first && !citation.NoteHasPrecedingText &&
            firstVisible != null && string.IsNullOrWhiteSpace(firstVisible.Prefix) && !firstVisible.AuthorOnly;
    }

    private sealed class Sequence {
        internal HashSet<string> Seen { get; } = new HashSet<string>(StringComparer.Ordinal);
        internal Dictionary<string, int> LastNotes { get; } = new Dictionary<string, int>(StringComparer.Ordinal);
        internal HashSet<string> NoteKeys { get; } = new HashSet<string>(StringComparer.Ordinal);
        internal CslCitation? Previous { get; set; }
    }
}
