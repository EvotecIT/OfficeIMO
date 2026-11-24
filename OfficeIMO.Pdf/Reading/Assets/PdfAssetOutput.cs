using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed class PdfAssetOutput {
    private const string OutputLimitMessage = "The PDF asset bundle exceeds its output byte limit.";
    private readonly Func<string, Stream> _create;
    private readonly CancellationToken _cancellationToken;
    private readonly List<string> _paths = new();
    private readonly Dictionary<(string Kind, string Hash), string> _byHash = new();
    private readonly Dictionary<string, int> _counts = new(StringComparer.Ordinal);
    private long _remainingBytes;
    private byte[]? _writeBuffer;

    internal PdfAssetOutput(Func<string, Stream> create, long maximumBytes, CancellationToken cancellationToken) {
        _create = create;
        _remainingBytes = maximumBytes;
        _cancellationToken = cancellationToken;
    }

    internal IReadOnlyList<string> Paths => _paths.AsReadOnly();

    internal string WriteDistinct(string kind, string extension, byte[] bytes, string hash) {
        if (_byHash.TryGetValue((kind, hash), out string? existing)) return existing;
        _counts.TryGetValue(kind, out int count);
        string path = kind + "/" + (count + 1).ToString("D4", CultureInfo.InvariantCulture) + "." + extension;
        Write(path, bytes);
        _counts[kind] = count + 1;
        _byHash.Add((kind, hash), path);
        return path;
    }

    internal void Write(string path, byte[] bytes) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (bytes.LongLength > _remainingBytes) throw PdfOutputLimitErrors.Create(OutputLimitMessage);
        using (Stream destination = _create(path) ?? throw new InvalidOperationException("The asset stream factory returned null.")) {
            // A destination stream can change the supplied byte array. Keep PDF-owned buffers private.
            byte[] buffer = _writeBuffer ??= new byte[64 * 1024];
            for (int offset = 0; offset < bytes.Length;) {
                _cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(buffer.Length, bytes.Length - offset);
                Buffer.BlockCopy(bytes, offset, buffer, 0, count);
                destination.Write(buffer, 0, count);
                offset += count;
            }
        }
        _remainingBytes -= bytes.LongLength;
        _paths.Add(path);
    }

    internal void WriteReport(string path, Action<TextWriter> write) {
        _cancellationToken.ThrowIfCancellationRequested();
        using (Stream destination = _create(path) ?? throw new InvalidOperationException("The asset stream factory returned null."))
        using (var bounded = new PdfBoundedWriteStream(destination, _remainingBytes, OutputLimitMessage)) {
            long start = bounded.Position;
            using (var writer = new StreamWriter(bounded, new UTF8Encoding(false), bufferSize: 4096, leaveOpen: true)) {
                write(writer);
            }
            _remainingBytes -= bounded.Position - start;
        }
        _paths.Add(path);
    }

    internal static IReadOnlyList<string> WriteDirectory(string outputDirectory,
        Func<Func<string, Stream>, IReadOnlyList<string>> write, Action<string>? validatePath = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNullOrWhiteSpace(outputDirectory, nameof(outputDirectory));
        cancellationToken.ThrowIfCancellationRequested();
        string destination = Path.GetFullPath(outputDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string? parent = Path.GetDirectoryName(destination);
        if (parent is null || Directory.Exists(destination) || File.Exists(destination)) {
            throw new IOException("The PDF asset output must be a new directory: " + destination);
        }

        validatePath?.Invoke(destination);
        string staging = Path.Combine(parent, ".pdf-assets-" + Guid.NewGuid().ToString("N"));
        validatePath?.Invoke(staging);
        Directory.CreateDirectory(staging);
        try {
            IReadOnlyList<string> paths = write(name => {
                string path = Path.Combine(staging, name.Replace('/', Path.DirectorySeparatorChar));
                validatePath?.Invoke(path);
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                return new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            });
            cancellationToken.ThrowIfCancellationRequested();
            validatePath?.Invoke(staging);
            validatePath?.Invoke(destination);
            Directory.Move(staging, destination);
            return paths;
        } finally {
            if (Directory.Exists(staging)) Directory.Delete(staging, recursive: true);
        }
    }
}
