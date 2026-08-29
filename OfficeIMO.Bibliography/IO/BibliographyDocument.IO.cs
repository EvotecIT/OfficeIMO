namespace OfficeIMO.Bibliography;

public sealed partial class BibliographyDocument {
    /// <summary>Loads bibliography data from a path, detecting format by extension and content when needed.</summary>
    public static BibliographyReadResult Load(string path, BibliographyFormat? format = null, BibliographyReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (format.HasValue) return Load(stream, format.Value, options, encoding, cancellationToken);
        if (BibliographyFormatDetector.TryDetectPath(path, out BibliographyFormat detected)) return Load(stream, detected, options, encoding, cancellationToken);
        return LoadDetected(stream, options, encoding, cancellationToken);
    }

    /// <summary>Loads bibliography data from a stream.</summary>
    public static BibliographyReadResult Load(Stream stream, BibliographyFormat format, BibliographyReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        options ??= new BibliographyReadOptions();
        options.Validate();
        byte[] bytes = ReadAllBytes(stream, options.MaximumInputBytes, cancellationToken);
        Encoding actualEncoding = encoding ?? BibliographyEncoding.Detect(bytes);
        string text = actualEncoding.GetString(BibliographyEncoding.RemovePreamble(bytes, actualEncoding));
        return BibliographyReader.Parse(text, format, options, bytes, cancellationToken);
    }

    /// <summary>Loads bibliography data from a path asynchronously.</summary>
    public static async Task<BibliographyReadResult> LoadAsync(string path, BibliographyFormat? format = null, BibliographyReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true);
        if (format.HasValue) return await LoadAsync(stream, format.Value, options, encoding, cancellationToken).ConfigureAwait(false);
        if (BibliographyFormatDetector.TryDetectPath(path, out BibliographyFormat detected)) return await LoadAsync(stream, detected, options, encoding, cancellationToken).ConfigureAwait(false);
        return await LoadDetectedAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Loads bibliography data from a stream asynchronously.</summary>
    public static async Task<BibliographyReadResult> LoadAsync(Stream stream, BibliographyFormat format, BibliographyReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        options ??= new BibliographyReadOptions();
        options.Validate();
        byte[] bytes = await ReadAllBytesAsync(stream, options.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
        Encoding actualEncoding = encoding ?? BibliographyEncoding.Detect(bytes);
        string text = actualEncoding.GetString(BibliographyEncoding.RemovePreamble(bytes, actualEncoding));
        return BibliographyReader.Parse(text, format, options, bytes, cancellationToken);
    }

    /// <summary>Saves or converts to a path.</summary>
    public BibliographyWriteResult Save(string path, BibliographyWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        BibliographyWriteResult result = Write(options, cancellationToken);
        File.WriteAllBytes(path, result.Bytes);
        return result;
    }

    /// <summary>Saves or converts to a stream without closing it.</summary>
    public BibliographyWriteResult Save(Stream stream, BibliographyWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        BibliographyWriteResult result = Write(options, cancellationToken);
        stream.Write(result.Bytes, 0, result.Bytes.Length);
        return result;
    }

    /// <summary>Saves or converts to a path asynchronously.</summary>
    public async Task<BibliographyWriteResult> SaveAsync(string path, BibliographyWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        BibliographyWriteResult result = Write(options, cancellationToken);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true);
        await stream.WriteAsync(result.Bytes, 0, result.Bytes.Length, cancellationToken).ConfigureAwait(false);
        return result;
    }

    /// <summary>Saves or converts to a stream asynchronously without closing it.</summary>
    public async Task<BibliographyWriteResult> SaveAsync(Stream stream, BibliographyWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        BibliographyWriteResult result = Write(options, cancellationToken);
        await stream.WriteAsync(result.Bytes, 0, result.Bytes.Length, cancellationToken).ConfigureAwait(false);
        return result;
    }

    private static byte[] ReadAllBytes(Stream stream, long maximum, CancellationToken cancellationToken) {
        var output = new MemoryStream();
        var buffer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            if (output.Length + read > maximum) throw new InvalidDataException($"Bibliography input exceeds the configured {maximum} byte limit.");
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static BibliographyReadResult LoadDetected(Stream stream, BibliographyReadOptions? options, Encoding? encoding, CancellationToken cancellationToken) {
        options ??= new BibliographyReadOptions(); options.Validate();
        byte[] bytes = ReadAllBytes(stream, options.MaximumInputBytes, cancellationToken); Encoding actualEncoding = encoding ?? BibliographyEncoding.Detect(bytes);
        string text = actualEncoding.GetString(BibliographyEncoding.RemovePreamble(bytes, actualEncoding));
        return BibliographyReader.Parse(text, BibliographyFormatDetector.Detect(text, options), options, bytes, cancellationToken);
    }

    private static async Task<BibliographyReadResult> LoadDetectedAsync(Stream stream, BibliographyReadOptions? options, Encoding? encoding, CancellationToken cancellationToken) {
        options ??= new BibliographyReadOptions(); options.Validate();
        byte[] bytes = await ReadAllBytesAsync(stream, options.MaximumInputBytes, cancellationToken).ConfigureAwait(false); Encoding actualEncoding = encoding ?? BibliographyEncoding.Detect(bytes);
        string text = actualEncoding.GetString(BibliographyEncoding.RemovePreamble(bytes, actualEncoding));
        return BibliographyReader.Parse(text, BibliographyFormatDetector.Detect(text, options), options, bytes, cancellationToken);
    }

    private static async Task<byte[]> ReadAllBytesAsync(Stream stream, long maximum, CancellationToken cancellationToken) {
        var output = new MemoryStream();
        var buffer = new byte[81920];
        while (true) {
            int read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            if (output.Length + read > maximum) throw new InvalidDataException($"Bibliography input exceeds the configured {maximum} byte limit.");
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

}
