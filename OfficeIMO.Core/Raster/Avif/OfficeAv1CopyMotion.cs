namespace OfficeIMO.Drawing;

/// <summary>One intra-block-copy displacement in eighth-luma-sample units.</summary>
internal readonly struct OfficeAv1CopyMotion {
    internal OfficeAv1CopyMotion(int row, int col) { Row=row; Col=col; }
    internal int Row { get; }
    internal int Col { get; }
}
