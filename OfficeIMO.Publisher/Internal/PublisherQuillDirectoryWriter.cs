namespace OfficeIMO.Publisher.Internal;

/// <summary>Retains source directory metadata and unknown padding while assigning complete descriptor rows across bounded directory pages.</summary>
internal sealed class PublisherQuillDirectoryWriter {
    private readonly PublisherQuillSourceLayout _layout;
    private readonly int _shift;
    private readonly List<(int Offset, int Count)> _pages;
    private PublisherQuillDirectoryWriter(PublisherQuillSourceLayout layout, int shift, List<(int Offset, int Count)> pages) {
        _layout = layout; _shift = shift; _pages = pages;
    }
    internal static PublisherQuillDirectoryWriter Plan(PublisherQuillSourceLayout layout, int count, int shift,
        Func<int> appendPage, CancellationToken token) {
        var pages = new List<(int Offset, int Count)>(); int remaining = count;
        foreach (PublisherQuillSourceDirectory directory in layout.Directories) {
            token.ThrowIfCancellationRequested(); if (remaining == 0) break;
            int capacity = directory.Count;
            for (int slot = capacity; slot < PublisherQuillSourceLayout.DescriptorsPerDirectory; slot++) {
                int position = directory.Offset + 8 + slot * 24;
                if (position + 24 > directory.PageOffset + 512 || Enumerable.Range(position, 24).Any(i => layout.Data.Bytes[i] != 0)) break;
                capacity++;
            }
            int assigned = Math.Min(capacity, remaining);
            pages.Add((directory.Offset == 24 ? 24 : directory.Offset + shift, assigned)); remaining -= assigned;
        }
        while (remaining > 0) {
            token.ThrowIfCancellationRequested(); int assigned = Math.Min(PublisherQuillSourceLayout.DescriptorsPerDirectory, remaining);
            pages.Add((appendPage(), assigned)); remaining -= assigned;
        }
        return new PublisherQuillDirectoryWriter(layout, shift, pages);
    }

    internal void Publish(byte[] output, IReadOnlyList<PublisherQuillOutputDescriptor> rows, CancellationToken token) {
        foreach (PublisherQuillSourceDirectory directory in _layout.Directories) {
            int offset = directory.Offset == 24 ? 24 : directory.Offset + _shift;
            Array.Clear(output, offset, 8 + directory.Count * 24);
        }
        int ordinal = 0;
        for (int page = 0; page < _pages.Count; page++) {
            token.ThrowIfCancellationRequested(); (int offset, int count) = _pages[page];
            Put16(output, offset, 504); Put16(output, offset + 2, count);
            Put(output, offset + 4, page + 1 < _pages.Count ? (uint)_pages[page + 1].Offset : uint.MaxValue);
            for (int i = 0; i < count; i++) {
                token.ThrowIfCancellationRequested(); PublisherQuillOutputDescriptor row = rows[ordinal++]; int target = offset + 8 + i * 24;
                Buffer.BlockCopy(_layout.Data.Bytes, _layout.DescriptorOffsets[row.Source], output, target, 24);
                Put16(output, target + 6, row.Id); Put(output, target + 16, (uint)row.Offset); Put(output, target + 20, (uint)row.Length);
            }
        }
        Put16(output, 12, rows.Count); Put(output, 16, (uint)(_pages.Count * 512)); Put(output, 20, (uint)output.Length);
    }

    private static void Put(byte[] bytes, int offset, uint value) { for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (i * 8)); }
    private static void Put16(byte[] bytes, int offset, int value) { bytes[offset] = (byte)value; bytes[offset + 1] = (byte)(value >> 8); }
}

internal readonly record struct PublisherQuillOutputDescriptor(PublisherQuillChunk Source, int Offset, int Length, int Id);
