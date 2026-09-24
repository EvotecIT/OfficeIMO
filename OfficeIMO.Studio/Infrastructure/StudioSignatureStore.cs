namespace OfficeIMO.Studio.Infrastructure;

internal enum StudioSignatureKind {
    Signature,
    Initials
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

    internal IReadOnlyList<StudioSavedSignature> List(StudioSignatureKind kind) {
        if (!Directory.Exists(_root)) return [];
        var result = new List<StudioSavedSignature>();
        foreach (FileInfo file in new DirectoryInfo(_root).EnumerateFiles(Prefix(kind) + "*.png").OrderByDescending(file => file.LastWriteTimeUtc)) {
            if (file.Length is <= 0 or > MaximumBytes) continue;
            try {
                RestrictExistingFile(file.FullName);
                RestrictExistingFile(System.IO.Path.ChangeExtension(file.FullName, ".json"));
                StudioSignatureShape? shape = ReadShape(file.FullName);
                result.Add(new StudioSavedSignature(kind, file.FullName, File.ReadAllBytes(file.FullName), shape?.Text,
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
        return new StudioSavedSignature(kind, path, png, text, strokes);
    }

    internal void Delete(StudioSavedSignature signature) {
        string full = System.IO.Path.GetFullPath(signature.Path);
        string root = System.IO.Path.GetFullPath(_root).TrimEnd(System.IO.Path.DirectorySeparatorChar) + System.IO.Path.DirectorySeparatorChar;
        StringComparison comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        if (!full.StartsWith(root, comparison)) throw new UnauthorizedAccessException("The saved signature is outside the signature folder.");
        File.Delete(System.IO.Path.ChangeExtension(full, ".json"));
        File.Delete(full);
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
