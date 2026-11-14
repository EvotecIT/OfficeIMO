using System.Globalization;
using System.Text;

namespace OfficeIMO.Archives;

/// <summary>Settings for <see cref="TarArchiveWriter"/>.</summary>
public sealed class TarArchiveWriterOptions : ArchiveWriterOptions {
    /// <summary>When true, the archive is wrapped in one gzip member (.tar.gz) with no file name and a zero MTIME.</summary>
    public bool GZip { get; set; }
}

/// <summary>Writes POSIX TAR archives: ustar headers, with pax extended headers for long or non-ASCII names, sizes of
/// 8 GiB or more and times outside the ustar range. With the default options the output depends only on the entries:
/// owner and group are 0 with empty names, modes are 0644 for files and 0755 for directories, and times are fixed.</summary>
/// <remarks>The archive ends with two zero blocks and is padded to a 10,240-byte record, as GNU tar and bsdtar do.
/// Member names pass the same path checks as extraction.</remarks>
public sealed class TarArchiveWriter : IDisposable {
    private const int Block = 512;
    private const int Record = 20 * Block;
    private const long OctalSizeLimit = 077777777777L;

    private readonly Stream _destination;
    private readonly Stream _output;
    private readonly bool _leaveOpen;
    private readonly TarArchiveWriterOptions _options;
    private readonly ArchivePaths _paths;
    private long _written;
    private bool _finished;
    private bool _faulted;
    private bool _disposed;

    /// <summary>Starts an archive at the destination's current position.</summary>
    public TarArchiveWriter(Stream destination, TarArchiveWriterOptions? options = null, bool leaveOpen = false) {
        ArgumentNullException.ThrowIfNull(destination);
        if (!destination.CanWrite) throw new ArgumentException("A writable destination stream is required.", nameof(destination));
        _destination = destination; _leaveOpen = leaveOpen;
        _options = options ?? new TarArchiveWriterOptions();
        _paths = _options.CreatePaths();
        _output = _options.GZip ? new GZipWriteStream(destination, new GZipWriterOptions(), leaveOpen: true) : destination;
    }

    /// <summary>Adds an explicit directory member.</summary>
    public void AddDirectory(string path, DateTimeOffset? lastWriteTime = null) {
        ThrowIfUnavailable();
        ArgumentNullException.ThrowIfNull(path);
        string normalized = path.Replace('\\', '/').TrimEnd('/');
        _paths.AddDirectory(normalized);
        Guard(() => WriteHeader(normalized + "/", '5', 0, lastWriteTime, 0x1ed));
    }

    /// <summary>Adds one file member from memory.</summary>
    public void AddFile(string path, byte[] content, DateTimeOffset? lastWriteTime = null) {
        ArgumentNullException.ThrowIfNull(content);
        using var source = new MemoryStream(content, writable: false);
        AddFile(path, source, lastWriteTime);
    }

    /// <summary>Adds one file member from the stream's current position to its end. A nonseekable stream is staged first
    /// so the header can carry its size. The stream is left open.</summary>
    public void AddFile(string path, Stream content, DateTimeOffset? lastWriteTime = null, CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        ArgumentNullException.ThrowIfNull(content);
        string normalized = ValidateFile(path);
        if (content.CanSeek) {
            long length = content.Length - content.Position;
            Guard(() => {
                WriteHeader(normalized, '0', length, lastWriteTime, 0x1a4);
                if (Copy(content, length, cancellationToken) != length || content.ReadByte() != -1)
                    throw new IOException("The source stream changed length while it was archived: " + path);
                Pad(length);
            });
            return;
        }
        using var staged = new ArchiveSpillStream();
        var buffer = new byte[81920];
        int read;
        while ((read = content.Read(buffer, 0, buffer.Length)) > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            staged.Write(buffer, 0, read);
        }
        staged.Position = 0;
        Guard(() => {
            WriteHeader(normalized, '0', staged.Length, lastWriteTime, 0x1a4);
            Copy(staged, staged.Length, cancellationToken);
            Pad(staged.Length);
        });
    }

    /// <summary>Adds one described entry.</summary>
    public void Add(ArchiveWriteEntry entry, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(entry);
        if (entry.IsDirectory) { AddDirectory(entry.Path, entry.LastWriteTime); return; }
        ThrowIfUnavailable();
        using Stream content = entry.OpenContent();
        AddFile(entry.Path, content, entry.LastWriteTime, cancellationToken);
    }

    /// <summary>Writes the end-of-archive blocks, record padding and any gzip trailer. Dispose also finishes.</summary>
    public void Finish() {
        if (_finished) return;
        ThrowIfUnavailable();
        _finished = true;
        long end = _written + 2 * Block;
        long padded = (end + Record - 1) / Record * Record;
        WriteZeros(padded - _written);
        if (_options.GZip) _output.Dispose();
        _destination.Flush();
    }

    /// <summary>Finishes the archive unless a member failed, then releases the destination unless it was left open.</summary>
    public void Dispose() {
        if (_disposed) return;
        try {
            if (!_faulted) Finish();
        } finally {
            _disposed = true;
            if (!_leaveOpen) _destination.Dispose();
        }
    }

    /// <summary>Writes the entries, in the given order, to a TAR stream.</summary>
    public static void Create(Stream destination, IEnumerable<ArchiveWriteEntry> entries, TarArchiveWriterOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(entries);
        using var writer = new TarArchiveWriter(destination, options, leaveOpen: true);
        foreach (ArchiveWriteEntry entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            writer.Add(entry, cancellationToken);
        }
        writer.Finish();
    }

    /// <summary>Writes the entries, in the given order, to a new TAR file. An existing file is not overwritten.</summary>
    public static void Create(string destinationPath, IEnumerable<ArchiveWriteEntry> entries, TarArchiveWriterOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(entries);
        ArchiveWriteSource.WriteToFile(destinationPath, output => Create(output, entries, options, cancellationToken));
    }

    /// <summary>Archives a folder's files and subfolders in ordinal path order. Symbolic links and reparse points are rejected.</summary>
    public static void CreateFromDirectory(string sourceDirectory, Stream destination, TarArchiveWriterOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new TarArchiveWriterOptions();
        Create(destination, ArchiveWriteSource.FromDirectory(sourceDirectory, options, cancellationToken), options, cancellationToken);
    }

    /// <summary>Archives a folder into a new TAR (or .tar.gz) file. An existing file is not overwritten.</summary>
    public static void CreateFromDirectory(string sourceDirectory, string destinationPath, TarArchiveWriterOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new TarArchiveWriterOptions();
        List<ArchiveWriteEntry> entries = ArchiveWriteSource.FromDirectory(sourceDirectory, options, cancellationToken);
        ArchiveWriteSource.WriteToFile(destinationPath, output => Create(output, entries, options, cancellationToken));
    }

    private string ValidateFile(string path) {
        ArgumentNullException.ThrowIfNull(path);
        if (path.EndsWith("/", StringComparison.Ordinal) || path.EndsWith("\\", StringComparison.Ordinal))
            throw new ArgumentException("File member names cannot end with a separator; use AddDirectory.", nameof(path));
        return _paths.Add(path);
    }

    private void Guard(Action write) {
        try { write(); } catch { _faulted = true; throw; }
    }

    private void WriteHeader(string name, char type, long size, DateTimeOffset? lastWriteTime, int mode) {
        long seconds = (_options.Resolve(lastWriteTime) ?? DateTimeOffset.FromUnixTimeSeconds(0)).ToUnixTimeSeconds();
        byte[] nameBytes = Encoding.UTF8.GetBytes(name);
        bool ascii = nameBytes.All(value => value < 128);
        bool split = TrySplit(nameBytes, out int prefixLength);
        var pax = new List<(string Key, string Value)>();
        if (!ascii || !split) pax.Add(("path", name));
        if (size > OctalSizeLimit) pax.Add(("size", size.ToString(CultureInfo.InvariantCulture)));
        if (seconds < 0 || seconds > OctalSizeLimit) pax.Add(("mtime", seconds.ToString(CultureInfo.InvariantCulture)));
        if (pax.Count > 0) {
            byte[] records = PaxRecords(pax);
            string baseName = name.TrimEnd('/');
            baseName = baseName.Substring(baseName.LastIndexOf('/') + 1);
            WriteBlock(Header(Placeholder("PaxHeaders/" + baseName), 'x', records.Length, seconds, 0x1a4, 0));
            Write(records, records.Length);
            Pad(records.Length);
        }
        if (ascii && split) WriteBlock(Header(nameBytes, type, size, seconds, mode, prefixLength));
        else WriteBlock(Header(Placeholder(name), type, size, seconds, mode, 0));
    }

    private static bool TrySplit(byte[] name, out int prefixLength) {
        prefixLength = 0;
        if (name.Length <= 100) return true;
        // ustar stores longer names as prefix (up to 155 bytes) + '/' + name (up to 100 bytes).
        for (int index = Math.Min(155, name.Length - 2); index > 0; index--) {
            if (name[index] != '/') continue;
            if (name.Length - index - 1 > 100) return false;
            prefixLength = index;
            return true;
        }
        return false;
    }

    private static byte[] Placeholder(string name) {
        // A deterministic ASCII stand-in for readers that ignore pax; pax-aware readers use the "path" record.
        var builder = new StringBuilder();
        foreach (char character in name) builder.Append(character < 128 && character >= 32 ? character : '_');
        string text = builder.ToString();
        if (text.Length > 100) text = text.Substring(text.Length - 100);
        return Encoding.ASCII.GetBytes(text);
    }

    private static byte[] PaxRecords(List<(string Key, string Value)> records) {
        var output = new MemoryStream();
        foreach (var (key, value) in records) {
            byte[] body = Encoding.UTF8.GetBytes(" " + key + "=" + value + "\n");
            int length = body.Length + 1;
            while (length.ToString(CultureInfo.InvariantCulture).Length + body.Length != length) length++;
            byte[] prefix = Encoding.ASCII.GetBytes(length.ToString(CultureInfo.InvariantCulture));
            output.Write(prefix, 0, prefix.Length);
            output.Write(body, 0, body.Length);
        }
        return output.ToArray();
    }

    private static byte[] Header(byte[] name, char type, long size, long seconds, int mode, int prefixLength) {
        var block = new byte[Block];
        if (prefixLength > 0) {
            Array.Copy(name, 0, block, 345, prefixLength);
            Array.Copy(name, prefixLength + 1, block, 0, name.Length - prefixLength - 1);
        } else Array.Copy(name, 0, block, 0, name.Length);
        Octal(block, 100, 8, mode);
        Octal(block, 108, 8, 0);
        Octal(block, 116, 8, 0);
        if (size > OctalSizeLimit) {
            // GNU base-256 for readers without pax; the pax "size" record is authoritative.
            block[124] = 0x80;
            for (int index = 0; index < 8; index++) block[135 - index] = (byte)(size >> (8 * index));
        } else Octal(block, 124, 12, size);
        Octal(block, 136, 12, seconds < 0 || seconds > OctalSizeLimit ? 0 : seconds);
        block[156] = (byte)type;
        Encoding.ASCII.GetBytes("ustar\0").CopyTo(block, 257);
        block[263] = (byte)'0'; block[264] = (byte)'0';
        Octal(block, 329, 8, 0);
        Octal(block, 337, 8, 0);
        for (int index = 148; index < 156; index++) block[index] = (byte)' ';
        long checksum = 0;
        foreach (byte value in block) checksum += value;
        Octal(block, 148, 7, checksum);
        block[155] = (byte)' ';
        return block;
    }

    private static void Octal(byte[] block, int offset, int length, long value) {
        string text = Convert.ToString(value, 8).PadLeft(length - 1, '0');
        for (int index = 0; index < length - 1; index++) block[offset + index] = (byte)text[index];
        block[offset + length - 1] = 0;
    }

    private long Copy(Stream source, long length, CancellationToken token) {
        var buffer = new byte[81920];
        long copied = 0;
        while (copied < length) {
            token.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, (int)Math.Min(buffer.Length, length - copied));
            if (read == 0) break;
            Write(buffer, read);
            copied += read;
        }
        return copied;
    }

    private void Pad(long length) {
        int remainder = (int)(length % Block);
        if (remainder != 0) WriteZeros(Block - remainder);
    }

    private void WriteZeros(long count) {
        var zeros = new byte[Block];
        while (count > 0) {
            int chunk = (int)Math.Min(count, Block);
            Write(zeros, chunk);
            count -= chunk;
        }
    }

    private void WriteBlock(byte[] block) => Write(block, block.Length);

    private void Write(byte[] buffer, int count) {
        _output.Write(buffer, 0, count);
        _written += count;
    }

    private void ThrowIfUnavailable() {
        if (_disposed) throw new ObjectDisposedException(nameof(TarArchiveWriter));
        if (_faulted) throw new InvalidOperationException("A previous member failed; the TAR archive is incomplete.");
        if (_finished) throw new InvalidOperationException("The TAR archive is already finished.");
    }
}
