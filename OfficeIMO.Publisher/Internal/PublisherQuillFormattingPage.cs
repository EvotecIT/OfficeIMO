namespace OfficeIMO.Publisher.Internal;

/// <summary>One native 512-byte formatting page, shared by preserved-source editing and template-free feasibility encoding.</summary>
internal static class PublisherQuillFormattingPage {
    internal const int PageBytes = 512;

    internal static byte[] Write(IReadOnlyList<PublisherQuillFormattingRange> ranges) {
        int count = ranges.Count;
        if (count is < 1 or > 84) throw new ArgumentException("The formatting page requires between 1 and 84 ranges.", nameof(ranges));
        return Encode(ranges, shared: false) ?? throw new InvalidDataException("The generated style exceeds its formatting page.");
    }

    /// <summary>Uses the observed native sharing of byte-identical OPLs within one page; null means the page is full.</summary>
    internal static byte[]? TryWriteShared(IReadOnlyList<PublisherQuillFormattingRange> ranges) =>
        ranges.Count is < 1 or > 84 ? null : Encode(ranges, shared: true);

    /// <summary>Reads the final absolute text-byte endpoint from a page produced by this codec.</summary>
    internal static uint LastEnd(byte[] page) => BitConverter.ToUInt32(page, 8 + (BitConverter.ToUInt16(page, 0) - 1) * 4);

    /// <summary>Links an encoded page to the preceding page's final text-byte endpoint.</summary>
    internal static void SetPreviousEnd(byte[] destination, int pageOffset, uint previousEnd) => Put(destination, pageOffset + 4, previousEnd);

    /// <summary>Writes native BTEC/BTEP page boundaries and absolute stream offsets into a planned destination.</summary>
    internal static void WriteIndex(byte[] destination, int target, IReadOnlyList<byte[]> pages, IReadOnlyList<int> offsets) {
        if (pages.Count == 0 || pages.Count != offsets.Count) throw new ArgumentException("Each formatting page requires one absolute offset.");
        Put(destination, target, (uint)pages.Count); Put(destination, target + 4, 4);
        Put(destination, target + 8, 0); Put(destination, target + 12, 0);
        for (int i = 0; i < pages.Count; i++) {
            Put(destination, target + 16 + i * 4, LastEnd(pages[i]));
            Put(destination, target + 16 + pages.Count * 4 + i * 4, (uint)offsets[i]);
        }
    }

    private static void Put(byte[] destination, int offset, uint value) {
        for (int i = 0; i < 4; i++) destination[offset + i] = (byte)(value >> (i * 8));
    }

    private static byte[]? Encode(IReadOnlyList<PublisherQuillFormattingRange> ranges, bool shared) {
        int count = ranges.Count;
        var styles = new List<byte[]>();
        var indexes = new List<int>();
        foreach (PublisherQuillFormattingRange range in ranges) {
            int index = shared ? styles.FindIndex(style => style.SequenceEqual(range.Style)) : -1;
            if (index < 0) { index = styles.Count; styles.Add(range.Style); }
            indexes.Add(index);
        }
        int headerBytes = 8 + 6 * count;
        int styleOffset = PageBytes - styles.Sum(style => style.Length);
        if (styleOffset < headerBytes) return null;
        int previousEnd = 0;
        foreach (PublisherQuillFormattingRange range in ranges) {
            if (range.AbsoluteEnd <= previousEnd) throw new ArgumentException("Formatting endpoints must increase.", nameof(ranges));
            previousEnd = range.AbsoluteEnd;
        }
        using var stream = new MemoryStream(PageBytes);
        using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, leaveOpen: true)) {
            writer.Write((ushort)count); writer.Write((ushort)1); writer.Write(0U);
            foreach (PublisherQuillFormattingRange range in ranges) writer.Write(range.AbsoluteEnd);
            var offsets = new List<int>();
            int offset = styleOffset;
            foreach (byte[] style in styles) { offsets.Add(offset); offset += style.Length; }
            foreach (int index in indexes) writer.Write((ushort)offsets[index]);
            writer.Write(new byte[styleOffset - headerBytes]);
            foreach (byte[] style in styles) writer.Write(style);
        }
        return stream.ToArray();
    }
}

internal sealed record PublisherQuillFormattingRange(int AbsoluteEnd, byte[] Style);
