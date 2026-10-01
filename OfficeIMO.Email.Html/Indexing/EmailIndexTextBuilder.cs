namespace OfficeIMO.Email;

/// <summary>Bounds normalized mail text while retaining classification offsets.</summary>
internal sealed class EmailIndexTextBuilder {
    private readonly int _maximum;
    private readonly StringBuilder _text = new StringBuilder();
    private readonly List<EmailIndexTextRegion> _regions = new List<EmailIndexTextRegion>();
    internal EmailIndexTextBuilder(int maximum) { _maximum = maximum; }
    internal bool Truncated { get; private set; }

    internal void Append(string text, EmailIndexRegionKind kind, string reason, bool preserveWhitespace) {
        for (int index = 0; index < text.Length && !Truncated; index++) {
            char value = text[index];
            if (!preserveWhitespace && char.IsWhiteSpace(value)) {
                if (_text.Length > 0 && !char.IsWhiteSpace(_text[_text.Length - 1])) Add(" ", kind, reason);
            } else if (value == '\r') {
                if (index + 1 < text.Length && text[index + 1] == '\n') index++;
                Add("\n", kind, reason);
            } else {
                int length = char.IsHighSurrogate(value) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]) ? 2 : 1;
                Add(text.Substring(index, length), kind, reason);
                index += length - 1;
            }
        }
    }
    internal void LineBreak(EmailIndexRegionKind kind, string reason) {
        if (_text.Length > 0 && _text[_text.Length - 1] != '\n') Add("\n", kind, reason);
    }
    private void Add(string value, EmailIndexRegionKind kind, string reason) {
        if (Truncated) return;
        if (value.Length > _maximum - _text.Length) { Truncated = true; return; }
        int offset = _text.Length;
        _text.Append(value);
        if (_regions.Count > 0 && _regions[_regions.Count - 1].Kind == kind && _regions[_regions.Count - 1].Reason == reason) {
            var last = _regions[_regions.Count - 1];
            _regions[_regions.Count - 1] = new EmailIndexTextRegion(last.Start, last.Length + value.Length, kind, reason);
        } else _regions.Add(new EmailIndexTextRegion(offset, value.Length, kind, reason));
    }
    internal EmailIndexTextResult Build(EmailBodySourceKind kind, bool excludeQuotes, bool excludeSignatures, IReadOnlyList<EmailDiagnostic> diagnostics) {
        string full = _text.ToString();
        var selected = new StringBuilder();
        bool excluded = false;
        foreach (var region in _regions) {
            if (excludeQuotes && region.Kind == EmailIndexRegionKind.Quoted || excludeSignatures && region.Kind == EmailIndexRegionKind.Signature) {
                excluded = true;
                continue;
            }
            if (excluded && selected.Length > 0 && !char.IsWhiteSpace(selected[selected.Length - 1]) && !char.IsWhiteSpace(full[region.Start])) selected.Append('\n');
            selected.Append(full, region.Start, region.Length);
            excluded = false;
        }
        return new EmailIndexTextResult(full, selected.ToString(), kind, _regions.AsReadOnly(), Truncated, diagnostics);
    }
}
