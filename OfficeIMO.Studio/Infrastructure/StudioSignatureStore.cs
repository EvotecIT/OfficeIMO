namespace OfficeIMO.Studio.Infrastructure;

internal enum StudioSignatureKind {
    Signature,
    Initials
}

internal static class StudioSignatureImage {
    private const long MaximumPreviewPixels = 4_000_000;

    internal static bool IsWithinPixelBudget(byte[] image) {
        if (OfficeIMO.Drawing.OfficeImageReader.TryIdentifyByContent(image, null, out OfficeIMO.Drawing.OfficeImageInfo info))
            return info.Width > 0 && info.Height > 0 && (long)info.Width * info.Height <= MaximumPreviewPixels;
        // Avalonia accepts some PNG/JPEG payloads whose ancillary data the stricter document
        // reader rejects. The container dimensions still bound decoding before Bitmap sees them.
        if (image.Length >= 24 && image.AsSpan(0, 8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }) &&
            image.AsSpan(12, 4).SequenceEqual("IHDR"u8)) {
            uint width = System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(image.AsSpan(16, 4));
            uint height = System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(image.AsSpan(20, 4));
            return width > 0 && height > 0 && width <= int.MaxValue && height <= int.MaxValue &&
                width <= MaximumPreviewPixels / height;
        }
        if (image.Length < 4 || image[0] != 0xff || image[1] != 0xd8) return false;
        for (int offset = 2; offset + 4 < image.Length;) {
            if (image[offset] != 0xff) return false;
            while (offset < image.Length && image[offset] == 0xff) offset++;
            if (offset + 2 >= image.Length) return false;
            byte marker = image[offset++];
            if (marker is 0xd8 or 0x01 || marker is >= 0xd0 and <= 0xd7) continue;
            int length = image[offset] << 8 | image[offset + 1];
            if (length < 2 || offset + length > image.Length) return false;
            if (marker is >= 0xc0 and <= 0xcf and not (0xc4 or 0xc8 or 0xcc)) {
                if (length < 7) return false;
                int height = image[offset + 3] << 8 | image[offset + 4];
                int width = image[offset + 5] << 8 | image[offset + 6];
                return width > 0 && height > 0 && (long)width * height <= MaximumPreviewPixels;
            }
            offset += length;
        }
        return false;
    }
}

internal sealed record StudioSavedSignature(StudioSignatureKind Kind, string Path, byte[] Png, string? Text = null,
    IReadOnlyList<IReadOnlyList<Avalonia.Point>>? Strokes = null);

internal sealed class StudioSignatureShape {
    public string? Text { get; set; }
    public List<double[]>? Strokes { get; set; }
}

/// <summary>
/// Keeps the user's reusable signature and initials images as PNG files in the Studio data folder.
/// Images never leave the machine; they are placed on pages as ordinary page content.
/// </summary>
internal sealed class StudioSignatureStore {
    internal const int MaximumPerKind = 6;
    private const int MaximumBytes = 4 * 1024 * 1024;
    private readonly string _root;

    internal StudioSignatureStore(string root) => _root = root ?? throw new ArgumentNullException(nameof(root));

    internal event Action? Changed;

    internal IReadOnlyList<StudioSavedSignature> List(StudioSignatureKind kind) {
        string[] paths;
        try { paths = Directory.GetFiles(_root, Prefix(kind) + "*.png"); }
        catch (DirectoryNotFoundException) { return []; }
        var result = new List<StudioSavedSignature>();
        var candidates = new List<(FileInfo File, DateTime Written)>();
        foreach (string path in paths) {
            try {
                var file = new FileInfo(path);
                if (file.Length is > 0 and <= MaximumBytes) candidates.Add((file, file.LastWriteTimeUtc));
            } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
        foreach ((FileInfo file, _) in candidates.OrderByDescending(candidate => candidate.Written)) {
            try {
                RestrictExistingFile(file.FullName);
                RestrictExistingFile(System.IO.Path.ChangeExtension(file.FullName, ".json"));
                StudioSignatureShape? shape = ReadShape(file.FullName);
                byte[] image = File.ReadAllBytes(file.FullName);
                if (!StudioSignatureImage.IsWithinPixelBudget(image)) continue;
                result.Add(new StudioSavedSignature(kind, file.FullName, image, shape?.Text,
                    shape?.Strokes?.Select(stroke => (IReadOnlyList<Avalonia.Point>)Enumerable.Range(0, stroke.Length / 2)
                        .Select(index => new Avalonia.Point(stroke[index * 2], stroke[index * 2 + 1])).ToArray()).ToArray()));
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
        return result;
    }

    internal StudioSavedSignature Save(StudioSignatureKind kind, byte[] png, string? text = null, IReadOnlyList<IReadOnlyList<Avalonia.Point>>? strokes = null) {
        ArgumentNullException.ThrowIfNull(png);
        if (png.Length is 0 or > MaximumBytes) throw new ArgumentException("The signature image is empty or too large.", nameof(png));
        if (!StudioSignatureImage.IsWithinPixelBudget(png)) throw new ArgumentException("The signature image exceeds the pixel limit or is invalid.", nameof(png));
        Directory.CreateDirectory(_root);
        string path = System.IO.Path.Combine(_root, Prefix(kind) + DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", System.Globalization.CultureInfo.InvariantCulture) + ".png");
        string sidecar = System.IO.Path.ChangeExtension(path, ".json");
        bool createdPng = false;
        try {
            WritePrivateFile(path, png);
            createdPng = true;
            if (text is not null || strokes is not null) {
                var shape = new StudioSignatureShape {
                    Text = text,
                    Strokes = strokes?.Select(stroke => stroke.SelectMany(point => new[] { point.X, point.Y }).ToArray()).ToList()
                };
                WritePrivateFile(sidecar, System.Text.Encoding.UTF8.GetBytes(System.Text.Json.JsonSerializer.Serialize(shape)));
            }
        } catch {
            // A failed sidecar write must not make an incomplete saved signature appear on restart.
            if (createdPng) {
                try { File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
            throw;
        }
        var saved = new StudioSavedSignature(kind, path, png, text, strokes);
        Changed?.Invoke();
        return saved;
    }

    internal void Delete(StudioSavedSignature signature) {
        string full = System.IO.Path.GetFullPath(signature.Path);
        string root = System.IO.Path.GetFullPath(_root).TrimEnd(System.IO.Path.DirectorySeparatorChar) + System.IO.Path.DirectorySeparatorChar;
        StringComparison comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        if (!full.StartsWith(root, comparison)) throw new UnauthorizedAccessException("The saved signature is outside the signature folder.");
        File.Delete(full);
        // Deleting the image first leaves the reusable shape intact if the image is locked.
        // If the sidecar cannot be removed, restore the image so the remembered signature is
        // still complete and usable after the next launch.
        try { File.Delete(System.IO.Path.ChangeExtension(full, ".json")); }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            try { WritePrivateFile(full, signature.Png); }
            catch (Exception restoreError) when (restoreError is IOException or UnauthorizedAccessException) {
                throw new AggregateException("The signature image could not be restored after sidecar deletion failed.", error, restoreError);
            }
            throw;
        }
        Changed?.Invoke();
    }

    private static void WritePrivateFile(string path, byte[] bytes) {
        var options = new FileStreamOptions { Mode = FileMode.CreateNew, Access = FileAccess.Write, Share = FileShare.None };
        if (!OperatingSystem.IsWindows()) options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
        using var stream = new FileStream(path, options);
        try { stream.Write(bytes); }
        catch {
            stream.Dispose();
            try { File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            throw;
        }
    }

    private static void RestrictExistingFile(string path) {
        if (!OperatingSystem.IsWindows() && File.Exists(path))
            File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite);
    }

    private static StudioSignatureShape? ReadShape(string pngPath) {
        string sidecar = System.IO.Path.ChangeExtension(pngPath, ".json");
        if (!File.Exists(sidecar) || new FileInfo(sidecar).Length > MaximumBytes) return null;
        try { return System.Text.Json.JsonSerializer.Deserialize<StudioSignatureShape>(File.ReadAllText(sidecar)); }
        catch (System.Text.Json.JsonException) { return null; }
    }

    private static string Prefix(StudioSignatureKind kind) => kind == StudioSignatureKind.Initials ? "initials-" : "signature-";
}
