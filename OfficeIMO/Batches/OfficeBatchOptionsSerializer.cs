using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.PowerPoint;

namespace OfficeIMO;

internal static class OfficeBatchOptionsSerializer {
    internal const int MaximumJobBytes = 1024 * 1024;
    private const string Schema = "officeimo.conversion.job.v2";
    private static readonly OfficeBatchDefinitionJsonContext Context = new(new JsonSerializerOptions {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow, MaxDepth = 32,
        Converters = {
            new OfficeBatchPageSizeConverter(),
            new OfficeBatchPageMarginsConverter(),
            new JsonStringEnumConverter<OfficeDocumentFormat>(allowIntegerValues: false),
            new JsonStringEnumConverter<OfficeConversionProfile>(allowIntegerValues: false),
            new JsonStringEnumConverter<OfficePublicationPolicy>(allowIntegerValues: false),
            new JsonStringEnumConverter<WordPdfImportMode>(allowIntegerValues: false),
            new JsonStringEnumConverter<PowerPointPdfImportMode>(allowIntegerValues: false),
            new JsonStringEnumConverter<ExcelPdfWorksheetLayoutMode>(allowIntegerValues: false),
            new JsonStringEnumConverter<PdfHtmlProfile>(allowIntegerValues: false)
        }
    });

    internal static string Serialize(OfficeBatchOptions options, string baseDirectory) {
        Validate(options);
        RequireAbsoluteBase(baseDirectory);
        string Relative(string path) => OfficeRuntimeCompatibility.GetRelativePath(baseDirectory, OfficeRuntimeCompatibility.GetFullPath(path, baseDirectory)).Replace(Path.DirectorySeparatorChar, '/');
        OfficeBatchOptions portable = options with {
            InputDirectory = options.InputDirectory == null ? null : Relative(options.InputDirectory),
            InputPaths = options.InputPaths?.Select(Relative).ToArray(), OutputDirectory = Relative(options.OutputDirectory),
            ResumeFrom = options.ResumeFrom == null ? null : Relative(options.ResumeFrom),
            ReportPath = options.ReportPath == null ? null : Relative(options.ReportPath)
        };
        using var stream = new BoundedReportStream(MaximumJobBytes);
        JsonSerializer.Serialize(stream, new OfficeBatchDefinition(Schema, portable, OfficeBatchPortableSettings.Capture(options.ConversionOptions)), Context.OfficeBatchDefinition);
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    internal static OfficeBatchOptions Deserialize(string json, string baseDirectory) {
        RequireAbsoluteBase(baseDirectory);
        if (Encoding.UTF8.GetByteCount(json) > MaximumJobBytes) throw new InvalidDataException("The saved batch exceeds its one MiB limit.");
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = 32 });
        OfficeBatchJson.RejectDuplicates(document.RootElement);
        OfficeBatchDefinition definition = JsonSerializer.Deserialize(json, Context.OfficeBatchDefinition)
            ?? throw new InvalidDataException("The saved batch is empty.");
        if (definition.Schema != Schema) throw new NotSupportedException("Unsupported saved-batch schema. Review and migrate the definition before running it.");
        if (definition.Batch == null || definition.Settings == null) throw new InvalidDataException("Batch selection and settings are required.");
        Validate(definition.Batch);
        string Resolve(string path) => OfficeRuntimeCompatibility.GetFullPath(path, baseDirectory);
        return definition.Batch with {
            InputDirectory = definition.Batch.InputDirectory == null ? null : Resolve(definition.Batch.InputDirectory),
            InputPaths = definition.Batch.InputPaths?.Select(Resolve).ToArray(), OutputDirectory = Resolve(definition.Batch.OutputDirectory),
            ResumeFrom = definition.Batch.ResumeFrom == null ? null : Resolve(definition.Batch.ResumeFrom),
            ReportPath = definition.Batch.ReportPath == null ? null : Resolve(definition.Batch.ReportPath),
            ConversionOptions = definition.Settings.CreateOptions()
        };
    }

    internal static OfficeBatchOptions ReadFile(string path) => ReadSnapshot(path, out _);
    internal static OfficeBatchOptions ReadSnapshot(string path, out string sha256) {
        string fullPath = Path.GetFullPath(path);
        OfficeBatchExecutor.EnsureNoLinks(fullPath);
        using var file = OfficePathIdentity.OpenRegularFileForRead(fullPath, Path.GetDirectoryName(fullPath)!, 81920);
        using var bytes = new BoundedReportStream(MaximumJobBytes);
        file.CopyTo(bytes);
        byte[] content = bytes.ToArray();
        string json = new UTF8Encoding(false, true).GetString(content).TrimStart('\uFEFF');
        OfficeBatchOptions options = Deserialize(json, Path.GetDirectoryName(fullPath)!);
        ValidateStoragePath(options, fullPath);
        sha256 = OfficeRuntimeCompatibility.ToHexString(OfficeRuntimeCompatibility.HashSha256(content));
        return options with { JobPath = fullPath, JobSha256 = sha256 };
    }

    internal static void SaveFile(OfficeBatchOptions options, string path, bool overwrite) {
        string fullPath = Path.GetFullPath(path);
        string json = Serialize(options, Path.GetDirectoryName(fullPath)!);
        ValidateStoragePath(options, fullPath);
        byte[] bytes = Encoding.UTF8.GetBytes(json);
        if (!overwrite) {
            OfficeFileCommit.WriteAtomically(fullPath, stream => { OfficeRuntimeCompatibility.WriteAll(stream, bytes); if (stream is FileStream file) file.Flush(true); },
                conflictPolicy: OfficeFileCommit.ConflictPolicy.FailIfExists);
            return;
        }
        _ = ReadSnapshot(fullPath, out string expected);
        OfficeFileCommit.WriteIfUnchangedAsync(fullPath, async (stream, token) => {
            await stream.WriteAsync(bytes, 0, bytes.Length, token).ConfigureAwait(false);
            if (stream is FileStream file) file.Flush(true);
        }, candidate => {
            try { _ = ReadSnapshot(candidate, out string actual); return expected == actual; }
            catch (Exception error) when (error is IOException or InvalidDataException or JsonException or NotSupportedException) { return false; }
        }).GetAwaiter().GetResult();
    }

    private static void Validate(OfficeBatchOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.InputDirectory == null && options.InputPaths is not { Length: > 0 }) throw new InvalidDataException("Select a source directory or explicit source files.");
        CheckText(options.Name, 256, false); CheckText(options.ConversionRouteId, 128, false);
        CheckText(options.InputDirectory, 4096, false); CheckText(options.OutputDirectory, 4096, true);
        CheckText(options.ResumeFrom, 4096, false); CheckText(options.ReportPath, 4096, false);
        if (options.InputPaths?.Length > 10_000) throw new InvalidDataException("Saved batches support at most 10,000 explicit files; use a source directory for larger selections.");
        foreach (string path in options.InputPaths ?? []) CheckText(path, 4096, true);
        if (options.Extensions?.Length > 256) throw new InvalidDataException("Too many source extension filters.");
        foreach (string extension in options.Extensions ?? []) CheckText(extension, 16, true);
        CheckText(options.TargetExtension, 16, true);
        if (options.MaxParallelism is < 1 or > 32 || options.MaximumFiles < 1 || options.MaximumInputBytes < 1 || options.MaximumOutputBytes < 1 || options.MaximumXmlCharactersInPart < 1 ||
            !OfficeRuntimeCompatibility.IsDefined(options.OutputProfile) || !OfficeRuntimeCompatibility.IsDefined(options.ConflictPolicy)) throw new InvalidDataException("Batch execution settings are invalid.");
        if (options.SourcePasswordEnvironmentVariable is { } reference) OfficeBatchJson.CheckEnvironmentName(reference);
        if (options.PdfPasswordEnvironmentVariable is { } pdfReference) OfficeBatchJson.CheckEnvironmentName(pdfReference);
    }
    internal static void ValidateStoragePath(OfficeBatchOptions options, string path) {
        Validate(options);
        string directory = Path.GetDirectoryName(Path.GetFullPath(path))!;
        OfficeBatchExecutor.EnsureNoLinks(path);
        foreach (string root in new[] { options.InputDirectory, options.OutputDirectory, options.CheckpointDirectory }.OfType<string>()) {
            string absolute = OfficeRuntimeCompatibility.GetFullPath(root, directory);
            OfficeBatchExecutor.EnsureNoLinks(absolute);
            if (OfficePathIdentity.IsSameOrDescendant(path, absolute)) throw new InvalidDataException("Save the batch outside its source, output and checkpoint trees.");
        }
        foreach (string input in options.InputPaths ?? [])
            if (OfficePathIdentity.IsSameOrDescendant(path, OfficeRuntimeCompatibility.GetFullPath(input, directory))) throw new InvalidDataException("The definition must not replace a source file or package member.");
    }
    private static void RequireAbsoluteBase(string directory) {
        if (!OfficeRuntimeCompatibility.IsPathFullyQualified(directory)) throw new ArgumentException("The batch base directory must be absolute.", nameof(directory));
    }
    private static void CheckText(string? value, int maximum, bool required) {
        if (required && string.IsNullOrWhiteSpace(value) || value?.Length > maximum || value?.Any(char.IsControl) == true)
            throw new InvalidDataException("A batch path or setting is empty, oversized or contains control characters.");
    }
}

internal sealed record OfficeBatchDefinition(string Schema, OfficeBatchOptions Batch, OfficeBatchPortableSettings Settings);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, GenerationMode = JsonSourceGenerationMode.Metadata,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(OfficeBatchDefinition))]
internal sealed partial class OfficeBatchDefinitionJsonContext : JsonSerializerContext;
