using OfficeIMO.Internal;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Security.Cryptography;
using OfficeIMO.Core.Internal;

namespace OfficeIMO;

/// <summary>Converts local files and captured streams through the owning OfficeIMO adapters.</summary>
public static partial class OfficeConverter {
    private static readonly ConcurrentDictionary<string, IOfficeConversionRoute> Registrations = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Registers an opt-in implementation of an existing catalog route. Built-in routes cannot be replaced.</summary>
    public static void Register(IOfficeConversionRoute route) {
        if (route == null) throw new ArgumentNullException(nameof(route));
        if (!OfficeConversionCapabilityCatalog.All.Any(capability => string.Equals(capability.Id, route.Id, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("Choose a route from the canonical conversion capability catalog.", nameof(route));
        if (BuiltinRouteIds.Contains(route.Id) || !Registrations.TryAdd(route.Id, route))
            throw new ArgumentException("The conversion route already has an implementation.", nameof(route));
    }

    /// <summary>Executable routes from the canonical capability catalog, including explicitly registered adapters.</summary>
    public static IReadOnlyList<OfficeConversionCapability> Routes => Array.AsReadOnly(OfficeConversionCapabilityCatalog.All
        .Where(capability => BuiltinRouteIds.Contains(capability.Id) || Registrations.ContainsKey(capability.Id)).ToArray());

    /// <summary>Converts a file and atomically publishes its completed output.</summary>
    public static OfficeConversionReport Convert(string source, string destination, OfficeConversionOptions? options = null) =>
        ConvertAsync(source, destination, options).GetAwaiter().GetResult();

    /// <summary>Converts a file, verifies its captured source and atomically publishes completed output.</summary>
    public static Task<OfficeConversionReport> ConvertAsync(string source, string destination,
        OfficeConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertFileAsync(source, destination, options, null, cancellationToken);

    internal static async Task<OfficeConversionReport> ConvertFileAsync(string source, string destination,
        OfficeConversionOptions? options, OfficeConversionExecutionInput? sourceOverride, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("A non-empty value is required.", nameof(source));
        if (string.IsNullOrWhiteSpace(destination)) throw new ArgumentException("A non-empty value is required.", nameof(destination));
        using var assets = new OfficeConversionAssets(cancellationToken);
        var timer = Stopwatch.StartNew();
        OfficeConversionOptions settings = (options ?? new()).Snapshot();
        string inputPath = Path.GetFullPath(source), outputPath = Path.GetFullPath(destination);
        if (Directory.Exists(inputPath)) inputPath = OfficeRuntimeCompatibility.TrimEndingDirectorySeparator(inputPath);
        if (OfficeIMO.Internal.OfficePathIdentity.AreEquivalent(inputPath, outputPath))
            throw new ArgumentException("The source and destination must be different files.", nameof(destination));
        settings.ResourceDirectory = Path.GetDirectoryName(inputPath);
        settings.ResourceSourcePath = inputPath;
        settings.SourceName = ConversionSourceName(inputPath, settings);
        settings.TargetName = Path.GetFileName(outputPath);
        if (sourceOverride == null && Directory.Exists(inputPath)) sourceOverride = CaptureDirectoryInput(inputPath, settings);
        if (sourceOverride?.SourcePublicationGuard is { } sourceGuard)
            settings.PublicationGuard = OfficePublicationGuards.Combine(settings.PublicationGuard, sourceGuard);
        using Stream input = sourceOverride == null
            ? OfficeConversionInputReader.OpenLocalRead(inputPath)
            : await sourceOverride.OpenRead(cancellationToken).ConfigureAwait(false);
        using OfficeBoundedMemoryStream captured = await CaptureAsync(input, settings.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
        byte[] sourceHash = OfficeRuntimeCompatibility.HashSha256(captured.GetBuffer(), 0, checked((int)captured.Length));
        settings.PublicationGuard = OfficePublicationGuards.Combine(settings.PublicationGuard,
            new CapturedSourceGuard(inputPath, sourceHash, settings.MaximumInputBytes, sourceOverride));
        if (sourceOverride?.ExpectedSha256 is { } expected && !string.Equals(expected, OfficeRuntimeCompatibility.ToHexString(sourceHash), StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The captured source changed before conversion.");
        using var output = new OfficeBoundedMemoryStream(settings.MaximumOutputBytes);
        OfficeConversionReport report = await ConvertCapturedAsync(captured, output, settings, cancellationToken).ConfigureAwait(false);
        if (settings.FailOnLoss) report.ThrowIfLoss();
        cancellationToken.ThrowIfCancellationRequested();
        try {
            byte[] currentHash;
            if (sourceOverride == null) {
                using Stream reopened = OfficeConversionInputReader.OpenLocalRead(inputPath);
                currentHash = await ComputeSourceHashAsync(reopened, settings.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
            } else {
                using Stream reopened = await sourceOverride.OpenRead(cancellationToken).ConfigureAwait(false);
                currentHash = await ComputeSourceHashAsync(reopened, settings.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
            }
            if (!OfficeRuntimeCompatibility.FixedTimeEquals(sourceHash, currentHash))
                throw new IOException("The source changed during conversion; no output was published.");
            await CheckStagingAsync(settings.PublicationGuard, Path.GetDirectoryName(outputPath)!, cancellationToken).ConfigureAwait(false);
            await CheckAssetsAsync(assets, settings, inputPath, outputPath, cancellationToken).ConfigureAwait(false);
            await assets.CompleteAsync((asset, token) => PublishAssetAsync(asset, settings, token)).ConfigureAwait(false);
            outputPath = await PublishConvertedOutputAsync(output, outputPath, settings, cancellationToken).ConfigureAwait(false);
        } catch (Exception error) when (error is not OperationCanceledException and not OutOfMemoryException and not StackOverflowException) {
            throw OfficeConversionFailures.Output(report, error);
        }
        return new OfficeConversionReport(inputPath, report.Target, outputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Converts captured input to a caller-owned stream. Streams remain open; loss rejection precedes destination writes.</summary>
    public static OfficeConversionReport Convert(Stream source, OfficeDocumentFormat sourceFormat, Stream destination,
        OfficeDocumentFormat target, OfficeConversionOptions? options = null) =>
        ConvertAsync(source, sourceFormat, destination, target, options).GetAwaiter().GetResult();

    /// <summary>Converts streams with bounded capture and output. A seekable source is read from its beginning and restored.</summary>
    public static async Task<OfficeConversionReport> ConvertAsync(Stream source, OfficeDocumentFormat sourceFormat,
        Stream destination, OfficeDocumentFormat target, OfficeConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        await ConvertStreamAsync(source, sourceFormat, destination, target, options, null, cancellationToken).ConfigureAwait(false);

    private static async Task<OfficeConversionReport> ConvertStreamAsync(Stream source, OfficeDocumentFormat sourceFormat,
        Stream destination, OfficeDocumentFormat target, OfficeConversionOptions? options, string? outputPath, CancellationToken cancellationToken) {
        if (source == null) throw new ArgumentNullException(nameof(source)); if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!source.CanRead || !destination.CanWrite) throw new ArgumentException("Readable input and writable output streams are required.");
        if (ReferenceEquals(source, destination)) throw new ArgumentException("Input and output streams must be different.");
        if (source is FileStream inputFile && destination is FileStream outputFile && IsSameFile(inputFile, outputFile))
            throw new ArgumentException("The source and destination must be different files.", nameof(destination));
        using var assets = new OfficeConversionAssets(cancellationToken);
        var timer = Stopwatch.StartNew();
        OfficeConversionOptions settings = (options ?? new()).Snapshot();
        settings.SourceName ??= "source" + (settings.SourceExtension ?? DefaultExtension(settings.SourceFormat ?? sourceFormat));
        settings.SourceName = ConversionSourceName(settings.SourceName, settings);
        settings.TargetName ??= "output" + DefaultExtension(target);
        if (FormatForExtension(Path.GetExtension(settings.SourceName)) != sourceFormat || FormatForExtension(Path.GetExtension(settings.TargetName)) != target)
            throw new ArgumentException("Stream filenames must agree with the supplied source and target formats.");
        long? position = source.CanSeek ? source.Position : null;
        try {
            if (source.CanSeek) source.Position = 0;
            using OfficeBoundedMemoryStream input = await CaptureAsync(source, settings.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
            using var output = new OfficeBoundedMemoryStream(settings.MaximumOutputBytes);
            OfficeConversionReport report = await ConvertCapturedAsync(input, output, settings, cancellationToken).ConfigureAwait(false);
            if (settings.FailOnLoss) report.ThrowIfLoss();
            cancellationToken.ThrowIfCancellationRequested();
            await OfficeConversionFailures.RunAsync(report, async () => {
                await CheckAssetsAsync(assets, settings, settings.ResourceSourcePath, outputPath, cancellationToken,
                    source as FileStream, destination as FileStream).ConfigureAwait(false);
                await assets.CompleteAsync((asset, token) => PublishAssetAsync(asset, settings, token)).ConfigureAwait(false);
                output.Position = 0;
                await output.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
            }).ConfigureAwait(false);
            return new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten, elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
        } finally { if (position.HasValue) source.Position = position.Value; }
    }

    /// <summary>Executes the conversion in private bounded storage and returns its findings without creating the destination.</summary>
    public static OfficeConversionReport Analyze(string source, string destination, OfficeConversionOptions? options = null) =>
        AnalyzeAsync(source, destination, options).GetAwaiter().GetResult();

    /// <summary>Assesses the actual decoding and serialization path without publishing an artifact.</summary>
    public static async Task<OfficeConversionReport> AnalyzeAsync(string source, string destination, OfficeConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        OfficeConversionOptions settings = (options ?? new()).Snapshot();
        if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("A source file is required.", nameof(source));
        if (string.IsNullOrWhiteSpace(destination)) throw new ArgumentException("A destination file is required.", nameof(destination));
        string inputPath = Path.GetFullPath(source);
        settings.SourceName = ConversionSourceName(inputPath, settings); settings.TargetName = Path.GetFileName(destination);
        settings.ResourceDirectory = Path.GetDirectoryName(inputPath);
        settings.ResourceSourcePath = inputPath;
        settings.FailOnLoss = false;
        using var input = OfficeConversionInputReader.OpenLocalRead(inputPath);
        using var assets = new OfficeConversionAssets(cancellationToken);
        using OfficeBoundedMemoryStream captured = await CaptureAsync(input, settings.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
        using var output = new OfficeBoundedMemoryStream(settings.MaximumOutputBytes);
        OfficeConversionReport report;
        if (IsNativePair(settings)) {
            _ = SelectRoute(settings);
            if (settings.OutputProfile != OfficeConversionProfile.Faithful) throw new ArgumentException("Native Office conversion supports the Faithful output profile.");
            report = await ConvertNativeAsync(captured, output, settings, cancellationToken, analyzeOnly: true).ConfigureAwait(false);
        } else {
            report = await ConvertCapturedAsync(captured, output, settings, cancellationToken).ConfigureAwait(false);
        }
        return new OfficeConversionReport(inputPath, report.Target, itemsWritten: report.ItemsWritten,
            elapsed: report.Elapsed, diagnostics: report.Diagnostics);
    }

    private static async Task<byte[]> ComputeSourceHashAsync(Stream source, long limit, CancellationToken token) {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = await source.ReadAsync(buffer, 0, buffer.Length, token).ConfigureAwait(false)) != 0) {
            total = checked(total + read);
            if (total > limit) throw new InvalidDataException("The verification source exceeds its input limit.");
            hash.AppendData(buffer, 0, read);
        }
        return hash.GetHashAndReset();
    }

    private static async Task<OfficeBoundedMemoryStream> CaptureAsync(Stream source, long limit, CancellationToken token) {
        var captured = new InputBuffer(limit);
        try { await source.CopyToAsync(captured, 81920, token).ConfigureAwait(false); captured.Position = 0; return captured; }
        catch { captured.Dispose(); throw; }
    }

    private sealed class InputBuffer(long maximumBytes) : OfficeBoundedMemoryStream(maximumBytes) {
        protected override Exception CreateLimitException(long limit) =>
            new InvalidDataException($"The input exceeds the configured input limit of {limit} bytes.");
    }

    private static Task CheckStagingAsync(IOfficePublicationGuard? guard, string directory, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        return guard is IOfficeStagingGuard staging ? staging.EnsureStagingDirectoryAllowedAsync(directory, token) : Task.CompletedTask;
    }

    internal static string DefaultExtension(OfficeDocumentFormat format) => format switch {
        OfficeDocumentFormat.Word => ".docx", OfficeDocumentFormat.Excel => ".xlsx", OfficeDocumentFormat.PowerPoint => ".pptx",
        OfficeDocumentFormat.Pdf => ".pdf", OfficeDocumentFormat.Markdown => ".md", OfficeDocumentFormat.Html => ".html",
        OfficeDocumentFormat.Rtf => ".rtf", OfficeDocumentFormat.Xps => ".xps", OfficeDocumentFormat.Publisher => ".pub",
        OfficeDocumentFormat.Text => ".txt", OfficeDocumentFormat.Email => ".eml", OfficeDocumentFormat.Epub => ".epub",
        OfficeDocumentFormat.Csv => ".csv", OfficeDocumentFormat.Json => ".json", OfficeDocumentFormat.Xml => ".xml",
        OfficeDocumentFormat.Visio => ".vsdx", _ => throw new NotSupportedException("The format requires an explicit physical filename and registered route.")
    };

    internal static OfficeDocumentFormat FormatForExtension(string? extension) => extension?.ToLowerInvariant() switch {
        ".doc" or ".docx" or ".dot" or ".dotx" or ".docm" or ".dotm" => OfficeDocumentFormat.Word,
        ".xls" or ".xla" or ".xlm" or ".xlw" or ".xlsx" or ".xlsb" or ".xlsm" or ".xlt" or ".xltx" or ".xltm" or ".xlam" => OfficeDocumentFormat.Excel,
        ".ppt" or ".ppa" or ".ppam" or ".pptx" or ".pot" or ".potx" or ".pps" or ".ppsx" or ".pptm" or ".potm" or ".ppsm" => OfficeDocumentFormat.PowerPoint,
        ".pdf" => OfficeDocumentFormat.Pdf, ".md" or ".markdown" => OfficeDocumentFormat.Markdown,
        ".html" or ".htm" => OfficeDocumentFormat.Html, ".txt" => OfficeDocumentFormat.Text,
        ".eml" or ".msg" or ".oft" or ".tnef" or ".dat" => OfficeDocumentFormat.Email,
        ".rtf" => OfficeDocumentFormat.Rtf, ".xps" or ".oxps" => OfficeDocumentFormat.Xps,
        ".pub" => OfficeDocumentFormat.Publisher, ".odg" or ".fodg" or ".odt" or ".ods" or ".odp" => OfficeDocumentFormat.OpenDocument,
        ".pages" or ".numbers" or ".key" => OfficeDocumentFormat.IWork,
        ".csv" or ".tsv" => OfficeDocumentFormat.Csv, ".json" => OfficeDocumentFormat.Json, ".xml" => OfficeDocumentFormat.Xml,
        ".epub" => OfficeDocumentFormat.Epub, ".vsdx" or ".vdx" or ".vtx" => OfficeDocumentFormat.Visio,
        _ => OfficeDocumentFormat.Unknown
    };
}
