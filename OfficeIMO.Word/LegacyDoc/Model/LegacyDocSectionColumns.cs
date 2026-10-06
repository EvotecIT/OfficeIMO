namespace OfficeIMO.Word.LegacyDoc.Model;

/// <summary>Reads the indexed column records and validates their complete native DOC layout.</summary>
internal sealed class LegacyDocSectionColumns {
    internal const ushort EvenlySpacedSprm = 0x3005;
    internal const ushort WidthSprm = 0xF203;
    internal const ushort SpacingSprm = 0xF204;
    internal const int MaximumCount = 44;
    internal const int MinimumWidth = 718;
    internal const int MaximumDimension = 32767;

    private bool _evenlySpaced = true;
    private readonly Dictionary<int, int> _widths = new();
    private readonly Dictionary<int, int> _spaces = new();
    private string? _warning;

    internal void InvalidCount() => _warning = "The native DOC section column count exceeds the supported binary format range.";

    internal bool TryRead(ushort sprm, byte[] bytes, int offset, int end, out int operandLength) {
        operandLength = sprm == EvenlySpacedSprm ? 1 : 3;
        if (sprm != EvenlySpacedSprm && sprm != WidthSprm && sprm != SpacingSprm) return false;
        if (end - offset < 2 + operandLength) {
            _warning = "The native DOC section contains a truncated column property.";
            operandLength = end - offset - 2;
            return true;
        }
        if (sprm == EvenlySpacedSprm) {
            _evenlySpaced = bytes[offset + 2] != 0;
            return true;
        }
        int index = bytes[offset + 2];
        int value = LegacyDocFib.ReadUInt16(bytes, offset + 3);
        if (index >= MaximumCount || value > MaximumDimension || (sprm == WidthSprm && value < MinimumWidth)) {
            _warning = "The native DOC section contains an invalid indexed column dimension.";
            return true;
        }
        (sprm == WidthSprm ? _widths : _spaces)[index] = value;
        return true;
    }

    internal IReadOnlyList<WordSectionColumn>? Build(int? columnCount, out string? warning) {
        int count = columnCount ?? 1;
        // Word retains indexed dimensions for inactive columns when the active count changes.
        // Only active columns need a width; cached dimensions still pass the operand checks above.
        if (!_evenlySpaced && _warning == null && Enumerable.Range(0, count).Any(index => !_widths.ContainsKey(index)))
            _warning = "The native DOC section does not contain a complete unequal-column layout.";
        warning = _warning;
        if (_evenlySpaced || warning != null) return null;
        return Array.AsReadOnly(Enumerable.Range(0, count).Select(index =>
            new WordSectionColumn(_widths[index], _spaces.TryGetValue(index, out int space) ? space : 0)).ToArray());
    }
}
