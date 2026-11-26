namespace OfficeIMO.Comic;

// Compare decimal runs by significant digits so zero padding and very long numbers cannot change numeric order.
internal sealed class ComicPageOrder : IComparer<string> {
    internal static readonly ComicPageOrder Instance = new();

    public int Compare(string? x, string? y) {
        if (x == null || y == null) return StringComparer.Ordinal.Compare(x, y);
        int a = 0, b = 0;
        while (a < x.Length && b < y.Length) {
            if (IsDigit(x[a]) && IsDigit(y[b])) {
                int endA = DigitEnd(x, a), endB = DigitEnd(y, b);
                while (a < endA - 1 && x[a] == '0') a++;
                while (b < endB - 1 && y[b] == '0') b++;
                int length = (endA - a).CompareTo(endB - b);
                if (length != 0) return length;
                int number = string.CompareOrdinal(x, a, y, b, endA - a);
                if (number != 0) return number;
                a = endA;
                b = endB;
            } else {
                int character = char.ToUpperInvariant(x[a++]).CompareTo(char.ToUpperInvariant(y[b++]));
                if (character != 0) return character;
            }
        }
        int remaining = (x.Length - a).CompareTo(y.Length - b);
        return remaining != 0 ? remaining : StringComparer.Ordinal.Compare(x, y);
    }

    private static bool IsDigit(char value) => value is >= '0' and <= '9';

    private static int DigitEnd(string value, int offset) {
        while (offset < value.Length && IsDigit(value[offset])) offset++;
        return offset;
    }
}
