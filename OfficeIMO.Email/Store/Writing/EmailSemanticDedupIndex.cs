using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email.Store;

/// <summary>Bounded-memory on-disk open-addressed set for fixed SHA-256 semantic fingerprints.</summary>
internal sealed class EmailSemanticDedupIndex : IDisposable {
    private const int DigestLength = 32;
    private const int RecordLength = DigestLength + 1;
    private readonly string _directory;
    private FileStream _stream;
    private readonly byte[] _record = new byte[RecordLength];
    private int _capacity;
    private int _count;
    private bool _disposed;

    internal EmailSemanticDedupIndex(string destinationPath, int initialCapacity = 1024) {
        if (initialCapacity <= 0) throw new ArgumentOutOfRangeException(nameof(initialCapacity));
        _capacity = NextPowerOfTwo(Math.Max(16, initialCapacity));
        string parent = Path.GetDirectoryName(Path.GetFullPath(destinationPath)) ?? Directory.GetCurrentDirectory();
        _directory = Path.Combine(parent,
            string.Concat(".OfficeIMO.Email.Store.Merge.", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(_directory);
        try {
            _stream = CreateFile(Path.Combine(_directory, "dedup.index"), _capacity);
        } catch {
            TryDeleteDirectory(_directory);
            throw;
        }
    }

    internal int Count => _count;
    internal long Length => _stream.Length;

    internal bool Contains(byte[] digest) {
        EnsureDigest(digest);
        return Find(digest, out _);
    }

    internal bool Add(byte[] digest) {
        EnsureDigest(digest);
        if ((_count + 1L) * 10L > _capacity * 7L) Grow();
        if (Find(digest, out long offset)) return false;
        _record[0] = 1;
        Buffer.BlockCopy(digest, 0, _record, 1, DigestLength);
        _stream.Position = offset;
        _stream.Write(_record, 0, _record.Length);
        _count++;
        return true;
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _stream.Dispose();
        TryDeleteDirectory(_directory);
    }

    private bool Find(byte[] digest, out long emptyOffset) {
        int mask = _capacity - 1;
        int slot = unchecked((int)Hash(digest)) & mask;
        for (int probe = 0; probe < _capacity; probe++) {
            long offset = checked((long)slot * RecordLength);
            _stream.Position = offset;
            ReadExactly(_stream, _record, 0, _record.Length);
            if (_record[0] == 0) {
                emptyOffset = offset;
                return false;
            }
            if (FixedTimeEquals(_record, digest)) {
                emptyOffset = offset;
                return true;
            }
            slot = (slot + 1) & mask;
        }
        throw new InvalidDataException("The semantic deduplication index has no empty slot.");
    }

    private void Grow() {
        if (_capacity > 1 << 27) throw new InvalidOperationException("The semantic deduplication index is too large.");
        int newCapacity = checked(_capacity * 2);
        string nextPath = Path.Combine(_directory, "dedup.next");
        try {
            using (FileStream next = CreateFile(nextPath, newCapacity)) {
                var record = new byte[RecordLength];
                _stream.Position = 0;
                for (int slot = 0; slot < _capacity; slot++) {
                    ReadExactly(_stream, record, 0, record.Length);
                    if (record[0] == 0) continue;
                    InsertInto(next, newCapacity, record);
                }
                next.Flush(flushToDisk: false);
            }
        } catch {
            TryDelete(nextPath);
            throw;
        }
        string currentPath = _stream.Name;
        _stream.Dispose();
        OfficeFileCommit.CommitTemporaryFile(nextPath, currentPath,
            OfficeFileCommit.ConflictPolicy.Replace);
        _stream = new FileStream(currentPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read,
            64 * 1024, FileOptions.RandomAccess);
        _capacity = newCapacity;
    }

    private static void InsertInto(FileStream stream, int capacity, byte[] occupiedRecord) {
        int mask = capacity - 1;
        var existing = new byte[RecordLength];
        int slot = unchecked((int)Hash(occupiedRecord, 1)) & mask;
        for (int probe = 0; probe < capacity; probe++) {
            long offset = checked((long)slot * RecordLength);
            stream.Position = offset;
            ReadExactly(stream, existing, 0, existing.Length);
            if (existing[0] == 0) {
                stream.Position = offset;
                stream.Write(occupiedRecord, 0, occupiedRecord.Length);
                return;
            }
            slot = (slot + 1) & mask;
        }
        throw new InvalidDataException("The grown semantic deduplication index has no empty slot.");
    }

    private static FileStream CreateFile(string path, int capacity) {
        var stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read,
            64 * 1024, FileOptions.RandomAccess);
        stream.SetLength(checked((long)capacity * RecordLength));
        return stream;
    }

    private static ulong Hash(byte[] digest, int offset = 0) {
        ulong value = 1469598103934665603UL;
        for (int index = 0; index < 8; index++) {
            value ^= digest[offset + index];
            value *= 1099511628211UL;
        }
        return value;
    }

    private static bool FixedTimeEquals(byte[] record, byte[] digest) {
        int difference = 0;
        for (int index = 0; index < DigestLength; index++) difference |= record[index + 1] ^ digest[index];
        return difference == 0;
    }

    private static void EnsureDigest(byte[] digest) {
        if (digest == null) throw new ArgumentNullException(nameof(digest));
        if (digest.Length != DigestLength) throw new ArgumentException("A SHA-256 digest is required.", nameof(digest));
    }

    private static int NextPowerOfTwo(int value) {
        int result = 1;
        while (result < value) result = checked(result * 2);
        return result;
    }

    private static void ReadExactly(Stream stream, byte[] buffer, int offset, int count) {
        int total = 0;
        while (total < count) {
            int read = stream.Read(buffer, offset + total, count - total);
            if (read == 0) throw new EndOfStreamException("The semantic deduplication index is truncated.");
            total += read;
        }
    }

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private static void TryDeleteDirectory(string path) {
        try { if (Directory.Exists(path)) Directory.Delete(path, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
