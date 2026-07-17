namespace OfficeIMO.Email.Store;

/// <summary>Disk-backed AMap payloads updated as allocations are made.</summary>
internal sealed class PstWriterAllocationMap : IDisposable {
    private const int PageDataLength = 496;
    private const long FirstAmapOffset = 0x4400;
    private const long AmapInterval = 0x3E000;
    private readonly string _path;
    private readonly FileStream _stream;
    private bool _disposed;

    internal PstWriterAllocationMap(string path) {
        _path = path;
        _stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.Read, 16 * 1024, FileOptions.RandomAccess);
    }

    internal void Mark(long allocationOffset, int allocationLength) {
        if (allocationLength <= 0) return;
        long allocationEnd = checked(allocationOffset + allocationLength);
        long firstIndex = Math.Max(0, (allocationOffset - FirstAmapOffset) / AmapInterval);
        for (long index = firstIndex; ; index++) {
            long coverageStart = FirstAmapOffset + index * AmapInterval;
            if (coverageStart >= allocationEnd) break;
            long coverageEnd = coverageStart + AmapInterval;
            long start = Math.Max(coverageStart, allocationOffset);
            long end = Math.Min(coverageEnd, allocationEnd);
            if (start >= end) continue;

            byte[] page = Read(index);
            int first = checked((int)((start - coverageStart) / 64));
            int last = checked((int)((end - 1 - coverageStart) / 64));
            for (int bit = first; bit <= last && bit < PageDataLength * 8; bit++) {
                page[bit / 8] |= checked((byte)(1 << (7 - bit % 8)));
            }
            Write(index, page);
        }
    }

    internal byte[] Read(long index) {
        var page = new byte[PageDataLength];
        long offset = checked(index * PageDataLength);
        if (offset >= _stream.Length) return page;
        _stream.Position = offset;
        int total = 0;
        while (total < page.Length) {
            int read = _stream.Read(page, total, page.Length - total);
            if (read == 0) break;
            total += read;
        }
        return page;
    }

    internal void Flush(bool durable) => _stream.Flush(durable);

    private void Write(long index, byte[] page) {
        _stream.Position = checked(index * PageDataLength);
        _stream.Write(page, 0, page.Length);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _stream.Dispose();
        TryDelete(_path);
    }

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
