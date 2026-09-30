using System;

namespace OfficeIMO.Drawing;

/// <summary>Normative AV1 transform dimensions and split order; values follow TX_SIZE numbering.</summary>
internal static class OfficeAv1TransformSize {
    private static readonly byte[] Widths = { 4,8,16,32,64,4,8,8,16,16,32,32,64,4,16,8,32,16,64 };
    private static readonly byte[] Heights = { 4,8,16,32,64,8,4,16,8,32,16,64,32,16,4,32,8,64,16 };
    private static readonly byte[] Splits = { 0,0,1,2,3,0,0,1,1,2,2,3,3,5,6,7,8,9,10 };
    internal static int Width(int size) { Validate(size); return Widths[size]; }
    internal static int Height(int size) { Validate(size); return Heights[size]; }
    internal static int Split(int size) { Validate(size); return Splits[size]; }
    internal static int Find(int width, int height) {
        for (int i=0; i<Widths.Length; i++) if (Widths[i]==width && Heights[i]==height) return i;
        throw new FormatException("Invalid AV1 transform dimensions.");
    }
    internal static int Maximum(int width, int height) => Find(Math.Min(64,width),Math.Min(64,height));
    internal static int Depth(int size) { int depth=0; while(size!=0) { size=Split(size); depth++; } return depth; }
    internal static int SquareUp(int size) => Log2(Math.Max(Width(size),Height(size))/4);
    private static int Log2(int value) { int n=0; while((value>>=1)!=0)n++; return n; }
    private static void Validate(int size) { if ((uint)size>=19) throw new ArgumentOutOfRangeException(nameof(size)); }
}

/// <summary>One residual transform in bitstream traversal order; origin is in that plane's pixel coordinates.</summary>
internal readonly struct OfficeAv1TransformBlock {
    internal OfficeAv1TransformBlock(int plane, int x, int y, int size) { Plane=plane; X=x; Y=y; Size=size; }
    internal int Plane { get; }
    internal int X { get; }
    internal int Y { get; }
    internal int Size { get; }
    internal int Width => OfficeAv1TransformSize.Width(Size);
    internal int Height => OfficeAv1TransformSize.Height(Size);
}

/// <summary>Immutable one-leaf transform grid and ordered residual geometry, without coefficient or type syntax.</summary>
internal sealed class OfficeAv1TransformLayout {
    private readonly byte[] _sizes;
    private readonly OfficeAv1TransformBlock[] _blocks;
    private readonly int _width;
    internal OfficeAv1TransformLayout(byte[] sizes, int width, int lastSize, OfficeAv1TransformBlock[] blocks) {
        _sizes=sizes; _width=width; LastSize=lastSize; _blocks=blocks;
    }
    internal int LastSize { get; }
    internal int Count => _blocks.Length;
    internal OfficeAv1TransformBlock Block(int index) => _blocks[index];
    internal int SizeAt(int row, int col) {
        if ((uint)col>=(uint)_width || (uint)row>=(uint)(_sizes.Length/_width)) throw new ArgumentOutOfRangeException();
        return _sizes[row*_width+col];
    }
}
