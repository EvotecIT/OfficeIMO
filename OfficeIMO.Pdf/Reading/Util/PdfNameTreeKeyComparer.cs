namespace OfficeIMO.Pdf;

/// <summary>Orders name-tree keys by their unsigned encoded bytes, including shorter prefixes first.</summary>
internal sealed class PdfNameTreeKeyComparer : IComparer<byte[]> {
    internal static PdfNameTreeKeyComparer Instance { get; } = new PdfNameTreeKeyComparer();

    public int Compare(byte[]? left, byte[]? right) {
        if (ReferenceEquals(left, right)) return 0;
        if (left is null) return -1;
        if (right is null) return 1;
        int count = Math.Min(left.Length, right.Length);
        for (int index = 0; index < count; index++) {
            int comparison = left[index].CompareTo(right[index]);
            if (comparison != 0) return comparison;
        }
        return left.Length.CompareTo(right.Length);
    }
}
