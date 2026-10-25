using System.Text;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Internal;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Strict, bounded JSON and safe local-file persistence for portable conversion jobs.</summary>
public static class OfficeConversionJobSerializer {
    /// <summary>Maximum UTF-8 job definition size, independent of document limits.</summary>
    public const int MaximumJobBytes = 1024 * 1024;
    private static readonly OfficeConversionJobJsonContext Context = new(new JsonSerializerOptions {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow, MaxDepth = 32,
        Converters = {
            new JsonStringEnumConverter<OfficeWorkflowOutputProfile>(allowIntegerValues: false),
            new JsonStringEnumConverter<OfficeWorkflowConflictPolicy>(allowIntegerValues: false),
            new JsonStringEnumConverter<OfficeConversionLossPolicy>(allowIntegerValues: false),
            new JsonStringEnumConverter<PdfWordImportMode>(allowIntegerValues: false),
            new JsonStringEnumConverter<PdfPowerPointImportMode>(allowIntegerValues: false),
            new JsonStringEnumConverter<ExcelPdfWorksheetLayoutMode>(allowIntegerValues: false),
            new JsonStringEnumConverter<PdfHtmlProfile>(allowIntegerValues: false)
        }
    });

    /// <summary>Serializes supported settings and secret references only.</summary>
    public static string Serialize(OfficeConversionJob job) {
        Validate(job);
        string json = JsonSerializer.Serialize(job, Context.OfficeConversionJob);
        if (Encoding.UTF8.GetByteCount(json) > MaximumJobBytes) throw new InvalidDataException("The saved job exceeds its one MiB limit.");
        return json;
    }

    /// <summary>Reads one schema version and rejects unknown fields, duplicate properties, numeric enums and oversized input.</summary>
    public static OfficeConversionJob Deserialize(string json) {
        ArgumentNullException.ThrowIfNull(json);
        if (Encoding.UTF8.GetByteCount(json) > MaximumJobBytes) throw new InvalidDataException("The saved job exceeds its one MiB limit.");
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = 32 });
        RejectDuplicates(document.RootElement);
        OfficeConversionJob job = JsonSerializer.Deserialize(json, Context.OfficeConversionJob)
            ?? throw new InvalidDataException("The saved job is empty.");
        Validate(job);
        return job;
    }

    /// <summary>Reads a bounded regular local file. The job must be outside its source, output and checkpoint trees.</summary>
    public static OfficeConversionJob ReadFile(string path) => ReadSnapshot(path, out _);

    private static OfficeConversionJob ReadSnapshot(string path, out string sha256) {
        string fullPath = Path.GetFullPath(path);
        OfficeConversionBatchExecutor.EnsureNoLinks(fullPath);
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > MaximumJobBytes) throw new InvalidDataException("The saved job exceeds its one MiB limit.");
        byte[] bytes = new byte[MaximumJobBytes + 1];
        int count = 0, read;
        while (count < bytes.Length && (read = stream.Read(bytes, count, bytes.Length - count)) > 0) count += read;
        if (count > MaximumJobBytes) throw new InvalidDataException("The saved job exceeds its one MiB limit.");
        string json = new UTF8Encoding(false, true).GetString(bytes, 0, count);
        OfficeConversionJob job = Deserialize(json.TrimStart('\uFEFF'));
        ValidateStoragePath(job, fullPath);
        sha256 = Convert.ToHexString(SHA256.HashData(bytes.AsSpan(0, count)));
        return job;
    }

    /// <summary>Saves atomically beside the destination. Explicit overwrite replaces only an unchanged valid job file; the parent must exist.</summary>
    public static void SaveFile(OfficeConversionJob job, string path, bool overwrite = false) {
        string json = Serialize(job);
        string fullPath = Path.GetFullPath(path);
        ValidateStoragePath(job, fullPath);
        string? expectedHash = null;
        if (File.Exists(fullPath)) {
            if (!overwrite) throw new IOException("The saved job already exists. Choose another path or explicitly overwrite it.");
            try { _ = ReadSnapshot(fullPath, out expectedHash); }
            catch (Exception error) when (error is JsonException or InvalidDataException or NotSupportedException) {
                throw new InvalidDataException("Only an existing valid conversion job may be replaced. Choose a separate job filename.", error);
            }
        }
        string temporary = fullPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                byte[] bytes = Encoding.UTF8.GetBytes(json);
                stream.Write(bytes);
                stream.Flush(flushToDisk: true);
            }
            OfficeConversionBatchExecutor.EnsureNoLinks(fullPath);
            if (expectedHash != null) {
                _ = ReadSnapshot(fullPath, out string currentHash);
                if (currentHash != expectedHash) throw new IOException("The saved job changed while preparing its replacement. Load it again before saving.");
            } else if (File.Exists(fullPath)) throw new IOException("The saved job destination became occupied. Choose another path or review the existing job.");
            File.Move(temporary, fullPath, expectedHash != null);
        } finally {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    internal static void Validate(OfficeConversionJob job) {
        ArgumentNullException.ThrowIfNull(job);
        if (job.Schema != OfficeConversionJob.CurrentSchema) throw new NotSupportedException("Unsupported conversion job schema.");
        if (job.Options == null || job.Secrets == null) throw new InvalidDataException("Job options and secret references are required.");
        if (job.InputDirectory == null && job.InputPaths is not { Length: > 0 }) throw new InvalidDataException("Select a source directory or explicit source files.");
        CheckText(job.Name, 256, "Job name"); CheckText(job.ConversionRouteId, 128, "Route");
        CheckPath(job.OutputDirectory); if (job.InputDirectory != null) CheckPath(job.InputDirectory);
        if (job.CheckpointDirectory != null) CheckPath(job.CheckpointDirectory);
        if (job.InputPaths?.Length > 10_000) throw new InvalidDataException("A saved job supports at most 10,000 explicit paths; use a source folder for larger jobs.");
        foreach (string path in job.InputPaths ?? []) CheckPath(path);
        if (job.SourceExtensions?.Length > 256) throw new InvalidDataException("Too many source extension filters.");
        foreach (string extension in job.SourceExtensions ?? []) CheckText(extension, 16, "Source extension", required: true);
        CheckText(job.TargetExtension, 16, "Target extension", required: true);
        CheckText(job.Options.PageRanges, 2048, "Page selection"); CheckText(job.Options.TextEncoding, 128, "Text encoding");
        if (job.MaximumConcurrency is < 1 or > 32 || job.MaximumFiles < 1 || job.MaximumInputBytes < 1 || job.MaximumOutputBytes < 1 || job.MaximumXmlCharactersInPart < 1)
            throw new InvalidDataException("Job execution limits are invalid.");
        if (job.Options.TabSize is < 1 or > 32 || job.Options.RasterDpi is { } dpi && (!double.IsFinite(dpi) || dpi is < 36 or > 600))
            throw new InvalidDataException("Job text or raster settings are invalid.");
        if (!Enum.IsDefined(job.OutputProfile) || !Enum.IsDefined(job.ConflictPolicy) || !Enum.IsDefined(job.Options.LegacyDocLossPolicy) ||
            job.Options.WordMode is { } word && !Enum.IsDefined(word) || job.Options.PowerPointMode is { } slides && !Enum.IsDefined(slides) ||
            job.Options.WorksheetLayout is { } layout && !Enum.IsDefined(layout) || job.Options.HtmlProfile is { } html && !Enum.IsDefined(html))
            throw new InvalidDataException("Choose supported job settings.");
        CheckEnvironmentName(job.Secrets.SourcePasswordEnvironmentVariable); CheckEnvironmentName(job.Secrets.PdfPasswordEnvironmentVariable);
    }

    private static void CheckPath(string? value) => CheckText(value, 4096, "Job path", required: true);
    private static void CheckText(string? value, int maximum, string label, bool required = false) {
        if (required && string.IsNullOrWhiteSpace(value) || value?.Length > maximum || value?.Any(char.IsControl) == true)
            throw new InvalidDataException(label + " is empty, oversized or contains control characters.");
    }
    private static void CheckEnvironmentName(string? name) {
        if (name == null) return;
        if (name.Length is < 1 or > 128 || !(char.IsAsciiLetter(name[0]) || name[0] == '_') ||
            name.Any(character => !(char.IsAsciiLetterOrDigit(character) || character == '_')))
            throw new InvalidDataException("Secret references must be environment-variable names.");
    }
    private static void RejectDuplicates(JsonElement element) {
        if (element.ValueKind == JsonValueKind.Object) {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (JsonProperty property in element.EnumerateObject()) {
                if (!names.Add(property.Name)) throw new JsonException("Duplicate job property: " + property.Name);
                RejectDuplicates(property.Value);
            }
        } else if (element.ValueKind == JsonValueKind.Array) foreach (JsonElement item in element.EnumerateArray()) RejectDuplicates(item);
    }
    /// <summary>Validates a job file's destination without writing. Hosts may use their own permission-aware publication after this check.</summary>
    public static void ValidateStoragePath(OfficeConversionJob job, string path) {
        Validate(job);
        path = Path.GetFullPath(path);
        OfficeConversionBatchExecutor.EnsureNoLinks(path);
        OfficeConversionBatchRequest request = job.CreateRequest(Path.GetDirectoryName(path)!);
        foreach (string root in new[] { request.InputDirectory, request.OutputDirectory, request.CheckpointDirectory }.OfType<string>()) {
            OfficeConversionBatchExecutor.EnsureNoLinks(root);
            if (OfficePathIdentity.IsSameOrDescendant(path, root)) throw new InvalidDataException("Save the job outside its source, output and checkpoint trees.");
        }
        foreach (string input in request.InputPaths ?? []) {
            OfficeConversionBatchExecutor.EnsureNoLinks(input);
            if (OfficePathIdentity.IsSameOrDescendant(path, input)) throw new InvalidDataException("A job file must not replace a source file or source package member.");
        }
    }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, GenerationMode = JsonSourceGenerationMode.Metadata,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(OfficeConversionJob))]
[JsonSerializable(typeof(OfficeConversionJobPlan))]
internal sealed partial class OfficeConversionJobJsonContext : JsonSerializerContext;
