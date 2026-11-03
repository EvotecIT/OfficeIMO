namespace OfficeIMO.Publisher.Internal;

/// <summary>Publishes bounded directory chains for fresh streams and retained source metadata without overwriting unknown padding.</summary>
internal sealed class PublisherQuillDirectoryWriter {
    private readonly PublisherQuillSourceLayout? _layout;
    private readonly int _shift;
    private readonly List<(int Offset, int Count)> _pages;
    private PublisherQuillDirectoryWriter(PublisherQuillSourceLayout? layout, int shift, List<(int Offset, int Count)> pages) {
        _layout = layout; _shift = shift; _pages = pages;
    }
    internal static PublisherQuillDirectoryWriter Plan(PublisherQuillSourceLayout layout, int count, int shift,
        Func<int> appendPage, CancellationToken token) {
        CheckCount(count);
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
        AppendPages(pages, remaining, appendPage, token);
        return new PublisherQuillDirectoryWriter(layout, shift, pages);
    }

    /// <summary>Keeps the first payload at 512 and appends any further directory pages through the stream owner's allocator.</summary>
    internal static PublisherQuillDirectoryWriter PlanFresh(int count, Func<int> appendPage, CancellationToken token) {
        token.ThrowIfCancellationRequested(); CheckCount(count);
        int first = Math.Min(PublisherQuillSourceLayout.DescriptorsPerDirectory, count);
        var pages = new List<(int Offset, int Count)> { (24, first) };
        AppendPages(pages, count - first, appendPage, token);
        return new PublisherQuillDirectoryWriter(null, 0, pages);
    }

    internal void Publish(byte[] output, IReadOnlyList<PublisherQuillOutputDescriptor> rows, CancellationToken token) {
        PublisherQuillSourceLayout layout = _layout ?? throw new InvalidOperationException("A fresh directory plan has no retained descriptors.");
        foreach (PublisherQuillSourceDirectory directory in layout.Directories) {
            token.ThrowIfCancellationRequested();
            int offset = directory.Offset == 24 ? 24 : directory.Offset + _shift;
            Array.Clear(output, offset, 8 + directory.Count * 24);
        }
        PublishRows(output, rows.Count, (ordinal, target) => {
            PublisherQuillOutputDescriptor row = rows[ordinal]; CheckRange(output, row.Offset, row.Length);
            Buffer.BlockCopy(layout.Data.Bytes, layout.DescriptorOffsets[row.Source], output, target, 24);
            Put16(output, target + 6, row.Id); Put(output, target + 16, (uint)row.Offset); Put(output, target + 20, (uint)row.Length);
        }, token);
    }

    /// <summary>Writes fresh descriptor metadata through the same chain publisher used by preserved-source output.</summary>
    internal void PublishFresh(byte[] output, IReadOnlyList<PublisherQuillFreshDescriptor> rows, CancellationToken token) {
        if (_layout != null) throw new InvalidOperationException("A retained directory plan requires its original descriptor templates.");
        PublishRows(output, rows.Count, (ordinal, target) => {
            PublisherQuillFreshDescriptor row = rows[ordinal]; CheckRange(output, row.Offset, row.Length);
            Put16(output, target, 24); Tag(row.Tag, target + 2); Put16(output, target + 6, row.Id);
            Put16(output, target + 8, 1); Put16(output, target + 10, 0); Tag(row.DescriptorTag, target + 12);
            Put(output, target + 16, (uint)row.Offset); Put(output, target + 20, (uint)row.Length);
        }, token);

        void Tag(string value, int offset) {
            if (value.Length != 4 || value.Any(character => character > 127)) throw new InvalidDataException("Quill descriptor tags require four ASCII bytes.");
            System.Text.Encoding.ASCII.GetBytes(value).CopyTo(output, offset);
        }
    }

    private void PublishRows(byte[] output, int count, Action<int, int> publishRow, CancellationToken token) {
        token.ThrowIfCancellationRequested(); CheckCount(count); CheckRange(output, 0, 32);
        if (_pages.Sum(page => page.Count) != count) throw new InvalidDataException("Quill descriptor count differs from its directory plan.");
        foreach (var page in _pages) CheckRange(output, page.Offset == 24 ? 0 : page.Offset, 512);
        int ordinal = 0;
        for (int page = 0; page < _pages.Count; page++) {
            token.ThrowIfCancellationRequested(); (int offset, int assigned) = _pages[page];
            Put16(output, offset, 504); Put16(output, offset + 2, assigned);
            Put(output, offset + 4, page + 1 < _pages.Count ? (uint)_pages[page + 1].Offset : uint.MaxValue);
            for (int i = 0; i < assigned; i++) {
                token.ThrowIfCancellationRequested(); publishRow(ordinal++, offset + 8 + i * 24);
            }
        }
        Put16(output, 12, count); Put(output, 16, checked((uint)(_pages.Count * 512))); Put(output, 20, (uint)output.Length);
    }

    private static void AppendPages(List<(int Offset, int Count)> pages, int remaining, Func<int> appendPage, CancellationToken token) {
        while (remaining > 0) {
            token.ThrowIfCancellationRequested(); int assigned = Math.Min(PublisherQuillSourceLayout.DescriptorsPerDirectory, remaining);
            int offset = appendPage();
            if (offset < 512 || offset % 512 != 0 || offset > int.MaxValue - 512 || pages.Any(page => page.Offset == offset))
                throw new InvalidDataException("Quill directory allocation requires distinct aligned pages.");
            pages.Add((offset, assigned)); remaining -= assigned;
        }
    }

    private static void CheckCount(int count) {
        if (count is < 1 or > PublisherQuillSourceLayout.MaximumDescriptors) throw new InvalidDataException("Quill descriptor limit exceeded.");
    }
    private static void CheckRange(byte[] output, int offset, int length) {
        if (offset < 0 || length < 0 || offset > output.Length - length) throw new InvalidDataException("Quill output range exceeds its stream.");
    }
    private static void Put(byte[] bytes, int offset, uint value) { for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (i * 8)); }
    private static void Put16(byte[] bytes, int offset, int value) {
        ushort encoded = checked((ushort)value); bytes[offset] = (byte)encoded; bytes[offset + 1] = (byte)(encoded >> 8);
    }
}

internal readonly record struct PublisherQuillOutputDescriptor(PublisherQuillChunk Source, int Offset, int Length, int Id);
internal readonly record struct PublisherQuillFreshDescriptor(string Tag, ushort Id, string DescriptorTag, int Offset, int Length);
