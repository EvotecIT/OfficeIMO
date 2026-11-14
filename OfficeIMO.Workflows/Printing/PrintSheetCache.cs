using System.IO.Compression;

namespace OfficeIMO.Workflows;

/// <summary>
/// Keeps encoded sheets while a transport repeats them for software copies, so each sheet is rasterized once per job.
/// The kept bytes stay within a capacity, normally the job's content byte limit; a sheet that does not fit is rendered again.
/// Sheets prepared ahead on worker threads add to it, so every access is synchronized.
/// </summary>
internal sealed class PrintSheetCache<T> where T : class {
    private readonly Dictionary<int, T> _sheets = new();
    private readonly long _capacity;
    private long _used;

    internal PrintSheetCache(long capacityBytes) => _capacity = Math.Max(0, capacityBytes);

    /// <summary>Times a kept sheet was used instead of rendering it again.</summary>
    internal int Reused { get { lock (_sheets) return _reused; } }
    private int _reused;

    internal bool TryGet(int key, out T value) {
        lock (_sheets) {
            if (_sheets.TryGetValue(key, out T? kept)) {
                _reused++;
                value = kept;
                return true;
            }
        }
        value = null!;
        return false;
    }

    /// <summary>Keeps <paramref name="value"/> when its <paramref name="bytes"/> fit the remaining capacity.</summary>
    internal void Add(int key, T value, long bytes) {
        lock (_sheets) {
            if (_sheets.ContainsKey(key) || bytes > _capacity - _used) return;
            _sheets.Add(key, value);
            _used += bytes;
        }
    }
}

/// <summary>Device pixels of one sheet, deflate-packed while they wait to be printed again.</summary>
internal sealed class PackedSheetPixels {
    private readonly byte[] _packed;

    private PackedSheetPixels(byte[] packed, int length, int width, int height, double dpi) {
        _packed = packed; Length = length; Width = width; Height = height; Dpi = dpi;
    }

    internal int Length { get; }
    internal int Width { get; }
    internal int Height { get; }
    internal double Dpi { get; }
    internal long PackedLength => _packed.LongLength;

    /// <summary>Packs pixels with fast deflate; printed sheets are mostly flat colour and pack to a small fraction of their size.</summary>
    internal static PackedSheetPixels Pack(byte[] pixels, int width, int height, double dpi) {
        using var packed = new MemoryStream();
        using (var deflate = new DeflateStream(packed, CompressionLevel.Fastest, leaveOpen: true)) deflate.Write(pixels, 0, pixels.Length);
        return new PackedSheetPixels(packed.ToArray(), pixels.Length, width, height, dpi);
    }

    internal byte[] Unpack() {
        var pixels = new byte[Length];
        using var deflate = new DeflateStream(new MemoryStream(_packed, writable: false), CompressionMode.Decompress);
        int offset = 0;
        while (offset < pixels.Length) {
            int read = deflate.Read(pixels, offset, pixels.Length - offset);
            if (read == 0) throw new InvalidDataException("A kept print sheet is incomplete.");
            offset += read;
        }
        return pixels;
    }
}
