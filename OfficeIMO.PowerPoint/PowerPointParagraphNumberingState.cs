namespace OfficeIMO.PowerPoint;

/// <summary>Evaluates paragraph markers in one text body for shared drawing and export consumers.</summary>
internal sealed class PowerPointParagraphNumberingState {
    private readonly Dictionary<int, int> _numbers = new();

    internal string? CreateMarker(PowerPointParagraph paragraph) {
        string? bullet = paragraph.BulletCharacter;
        if (!string.IsNullOrEmpty(bullet)) return bullet + " ";
        if (!paragraph.IsNumbered) return null;
        int level = paragraph.Level ?? 0;
        // A parent item starts a new child sequence, including after an explicit
        // child restart. Sibling and lower-level items preserve their parents.
        foreach (int deeper in _numbers.Keys.Where(item => item > level).ToArray())
            _numbers.Remove(deeper);
        int current = paragraph.NumberingStartAt
            ?? (_numbers.TryGetValue(level, out int previous) ? checked(previous + 1) : 1);
        _numbers[level] = current;
        return PowerPointNumberingFormatter.FormatMarker(current, paragraph.NumberingScheme) + " ";
    }
}
