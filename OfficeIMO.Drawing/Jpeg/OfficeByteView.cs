using System;

namespace OfficeIMO.Drawing;

/// <summary>Small allocation-free byte-array view used by the managed JPEG codec on all target frameworks.</summary>
internal readonly struct OfficeByteView {
    private readonly byte[] _data;
    private readonly int _offset;

    public OfficeByteView(byte[] data) : this(data, 0, data?.Length ?? 0) { }

    private OfficeByteView(byte[] data, int offset, int length) {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        if (offset < 0 || length < 0 || offset > data.Length - length) throw new ArgumentOutOfRangeException(nameof(offset));
        _offset = offset;
        Length = length;
    }

    public int Length { get; }

    public byte this[int index] {
        get {
            if ((uint)index >= (uint)Length) throw new ArgumentOutOfRangeException(nameof(index));
            return _data[_offset + index];
        }
    }

    public OfficeByteView Slice(int start) => Slice(start, Length - start);

    public OfficeByteView Slice(int start, int length) {
        if (start < 0 || length < 0 || start > Length - length) throw new ArgumentOutOfRangeException(nameof(start));
        return new OfficeByteView(_data, _offset + start, length);
    }

    public byte[] ToArray() {
        byte[] copy = new byte[Length];
        Buffer.BlockCopy(_data, _offset, copy, 0, Length);
        return copy;
    }

    public static implicit operator OfficeByteView(byte[] data) => new OfficeByteView(data);
}

internal static class OfficeByteViewExtensions {
    public static OfficeByteView Slice(this byte[] data, int start) => new OfficeByteView(data).Slice(start);
    public static OfficeByteView Slice(this byte[] data, int start, int length) => new OfficeByteView(data).Slice(start, length);
}
