using System.Collections;
using System.Globalization;
using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

/// <summary>Writes completed indirect objects once while retaining only xref offsets and materialization state.</summary>
internal sealed class PdfForwardOnlyObjectStore : IPdfObjectStore {
    private readonly Stream _destination;
    private readonly HashAlgorithm _fileIdHash;
    private readonly List<long> _offsets = new();
    private readonly List<bool> _materialized = new();
    private long _written;
    private long _largestSerializedObjectBytes;
    private bool _completed;
    private bool _disposed;

    internal PdfForwardOnlyObjectStore(Stream destination, PdfFileVersion fileVersion) {
        Guard.NotNull(destination, nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        _destination = destination;
        _fileIdHash = SHA256.Create();
        byte[] header = PdfEncoding.Latin1GetBytes(
            "%PDF-" + PdfFileAssembler.GetHeaderVersion(fileVersion) + "\n%\u00e2\u00e3\u00cf\u00d3\n");
        WriteSegment(header);
    }

    internal long LargestSerializedObjectBytes => _largestSerializedObjectBytes;
    internal long BytesWritten => _written;
    public int Count => _materialized.Count;
    public bool IsReadOnly => false;

    public byte[] this[int index] {
        get => throw new NotSupportedException("Forward-only PDF objects cannot be read after emission.");
        set {
            ThrowIfUnavailable();
            Guard.NotNull(value, nameof(value));
            if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
            if (_materialized[index]) {
                throw new InvalidOperationException("A forward-only PDF object cannot be emitted more than once.");
            }
            WriteObject(index, new[] { value }, value.LongLength);
        }
    }

    internal int Reserve() {
        ThrowIfUnavailable();
        _offsets.Add(0L);
        _materialized.Add(false);
        return Count;
    }

    public void Add(byte[] item) {
        ThrowIfUnavailable();
        Guard.NotNull(item, nameof(item));
        int index = Count;
        _offsets.Add(0L);
        _materialized.Add(false);
        WriteObject(index, new[] { item }, item.LongLength);
    }

    internal void AddSegments(params byte[][] segments) {
        ThrowIfUnavailable();
        Guard.NotNull(segments, nameof(segments));
        long length = 0L;
        for (int index = 0; index < segments.Length; index++) {
            Guard.NotNull(segments[index], nameof(segments));
            length = AddWithoutOverflow(length, segments[index].LongLength);
        }

        int objectIndex = Count;
        _offsets.Add(0L);
        _materialized.Add(false);
        WriteObject(objectIndex, segments, length);
    }

    internal long Complete(int catalogId, int infoId, string? trailerIdEntry = null) {
        ThrowIfUnavailable();
        for (int index = 0; index < _materialized.Count; index++) {
            if (!_materialized[index]) {
                throw new InvalidOperationException(
                    "PDF object " + (index + 1).ToString(CultureInfo.InvariantCulture) +
                    " was reserved but never materialized.");
            }
        }

        _fileIdHash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        byte[] fullHash = _fileIdHash.Hash ?? throw new InvalidOperationException("Unable to calculate the PDF trailer file identifier.");
        var fileId = new byte[16];
        Buffer.BlockCopy(fullHash, 0, fileId, 0, fileId.Length);

        long xrefPosition = _written;
        var trailer = new System.Text.StringBuilder();
        trailer.Append("xref\n0 ").Append((Count + 1).ToString(CultureInfo.InvariantCulture)).Append('\n');
        trailer.Append("0000000000 65535 f \n");
        for (int index = 0; index < _offsets.Count; index++) {
            trailer.Append(_offsets[index].ToString("0000000000", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }

        string idEntry = string.IsNullOrWhiteSpace(trailerIdEntry)
            ? " /ID [" + PdfSyntaxEscaper.HexString(fileId) + " " + PdfSyntaxEscaper.HexString(fileId) + "]"
            : trailerIdEntry!;
        trailer.Append("trailer\n<< /Size ").Append((Count + 1).ToString(CultureInfo.InvariantCulture))
            .Append(" /Root ").Append(PdfSyntaxEscaper.IndirectReference(catalogId))
            .Append(infoId > 0 ? " /Info " + PdfSyntaxEscaper.IndirectReference(infoId) : string.Empty)
            .Append(idEntry).Append(" >>\n")
            .Append("startxref\n").Append(xrefPosition.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        byte[] trailerBytes = System.Text.Encoding.ASCII.GetBytes(trailer.ToString());
        _destination.Write(trailerBytes, 0, trailerBytes.Length);
        _written += trailerBytes.LongLength;
        _completed = true;
        return _written;
    }

    public void Clear() => throw new NotSupportedException("Forward-only PDF objects cannot be cleared.");
    public bool Contains(byte[] item) => false;
    public void CopyTo(byte[][] array, int arrayIndex) => throw new NotSupportedException("Forward-only PDF objects cannot be copied.");
    public IEnumerator<byte[]> GetEnumerator() => throw new NotSupportedException("Forward-only PDF objects cannot be enumerated.");
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    public int IndexOf(byte[] item) => -1;
    public void Insert(int index, byte[] item) => throw new NotSupportedException("Forward-only PDF objects are append-only.");
    public bool Remove(byte[] item) => throw new NotSupportedException("Forward-only PDF objects cannot be removed.");
    public void RemoveAt(int index) => throw new NotSupportedException("Forward-only PDF objects cannot be removed.");

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _fileIdHash.Dispose();
        _offsets.Clear();
        _materialized.Clear();
    }

    private void WriteObject(int index, byte[][] segments, long length) {
        _offsets[index] = _written;
        for (int segmentIndex = 0; segmentIndex < segments.Length; segmentIndex++) {
            WriteSegment(segments[segmentIndex]);
        }
        _materialized[index] = true;
        _largestSerializedObjectBytes = Math.Max(_largestSerializedObjectBytes, length);
    }

    private void WriteSegment(byte[] bytes) {
        _fileIdHash.TransformBlock(bytes, 0, bytes.Length, bytes, 0);
        _destination.Write(bytes, 0, bytes.Length);
        _written += bytes.LongLength;
    }

    #pragma warning disable CA1513 // Newer helper is unavailable on netstandard2.0 and net472.
    private void ThrowIfUnavailable() {
        if (_disposed) throw new ObjectDisposedException(nameof(PdfForwardOnlyObjectStore));
        if (_completed) throw new InvalidOperationException("Forward-only PDF serialization is already complete.");
    }
    #pragma warning restore CA1513

    private static long AddWithoutOverflow(long left, long right) =>
        left > long.MaxValue - right ? long.MaxValue : left + right;
}
