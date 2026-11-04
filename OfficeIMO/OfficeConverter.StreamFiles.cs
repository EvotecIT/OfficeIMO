using System.Diagnostics;
using OfficeIMO.Core.Internal;

namespace OfficeIMO;

public static partial class OfficeConverter {
    /// <summary>Converts a caller-owned stream to a file using default conversion options.</summary>
    public static OfficeConversionReport Convert(Stream source, OfficeDocumentFormat sourceFormat, string destination) =>
        Convert(source, sourceFormat, destination, null);

    /// <summary>Converts a caller-owned stream and atomically publishes a file selected by its physical extension.</summary>
    public static OfficeConversionReport Convert(Stream source, OfficeDocumentFormat sourceFormat, string destination,
        OfficeConversionOptions? options) => ConvertAsync(source, sourceFormat, destination, options, default).GetAwaiter().GetResult();

    /// <summary>Converts a caller-owned stream to a file using default conversion options.</summary>
    public static Task<OfficeConversionReport> ConvertAsync(Stream source, OfficeDocumentFormat sourceFormat, string destination) =>
        ConvertAsync(source, sourceFormat, destination, null, default);

    /// <summary>Converts a caller-owned stream to a file using the supplied conversion options.</summary>
    public static Task<OfficeConversionReport> ConvertAsync(Stream source, OfficeDocumentFormat sourceFormat, string destination,
        OfficeConversionOptions? options) => ConvertAsync(source, sourceFormat, destination, options, default);

    /// <summary>Converts a caller-owned stream with bounded output, loss policy and atomic file publication. The source remains open and its seekable position is restored.</summary>
    public static async Task<OfficeConversionReport> ConvertAsync(Stream source, OfficeDocumentFormat sourceFormat, string destination,
        OfficeConversionOptions? options, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(destination)) throw new ArgumentException("A destination file is required.", nameof(destination));
        string outputPath = Path.GetFullPath(destination);
        if (source is FileStream file && IsSameFile(file, outputPath))
            throw new ArgumentException("The source and destination must be different files.", nameof(destination));
        var timer = Stopwatch.StartNew();
        var settings = (options ?? new()).Snapshot();
        settings.TargetName = Path.GetFileName(outputPath);
        using var output = new OfficeBoundedMemoryStream(settings.MaximumOutputBytes);
        var report = await ConvertStreamAsync(source, sourceFormat, output, FormatForExtension(Path.GetExtension(outputPath)),
            settings, outputPath, cancellationToken).ConfigureAwait(false);
        try {
            outputPath = await PublishConvertedOutputAsync(output, outputPath, settings, cancellationToken).ConfigureAwait(false);
        } catch (Exception error) when (error is not OperationCanceledException and not OutOfMemoryException and not StackOverflowException) {
            throw OfficeConversionFailures.Output(report, error);
        }
        return new OfficeConversionReport(report.Source, report.Target, outputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    // FileStream.Name is not available for streams created from native handles. Inspect the
    // open file itself so aliases and hard links cannot bypass source/output protection.
    private static bool IsSameFile(FileStream stream, string path) =>
        OfficePathIdentity.SupportsPhysicalIdentity
            ? File.Exists(path) && string.Equals(OfficePathIdentity.GetPhysicalIdentityKey(stream.Name, stream.SafeFileHandle),
                OfficePathIdentity.GetPhysicalIdentityKey(path), StringComparison.Ordinal)
            : File.Exists(stream.Name) && OfficePathIdentity.AreEquivalent(stream.Name, path);

    private static bool IsSameFile(FileStream source, FileStream destination) =>
        OfficePathIdentity.SupportsPhysicalIdentity
            ? string.Equals(OfficePathIdentity.GetPhysicalIdentityKey(source.Name, source.SafeFileHandle),
                OfficePathIdentity.GetPhysicalIdentityKey(destination.Name, destination.SafeFileHandle), StringComparison.Ordinal)
            : IsSameFile(source, destination.Name);

    /// <summary>Converts a local file to a caller-owned stream using default conversion options.</summary>
    public static OfficeConversionReport Convert(string source, Stream destination, OfficeDocumentFormat target) =>
        Convert(source, destination, target, null);

    /// <summary>Converts a local file to a caller-owned stream, retaining its source-relative resource policy.</summary>
    public static OfficeConversionReport Convert(string source, Stream destination, OfficeDocumentFormat target,
        OfficeConversionOptions? options) => ConvertAsync(source, destination, target, options, default).GetAwaiter().GetResult();

    /// <summary>Converts a local file to a caller-owned stream using default conversion options.</summary>
    public static Task<OfficeConversionReport> ConvertAsync(string source, Stream destination, OfficeDocumentFormat target) =>
        ConvertAsync(source, destination, target, null, default);

    /// <summary>Converts a local file to a caller-owned stream using the supplied conversion options.</summary>
    public static Task<OfficeConversionReport> ConvertAsync(string source, Stream destination, OfficeDocumentFormat target,
        OfficeConversionOptions? options) => ConvertAsync(source, destination, target, options, default);

    /// <summary>Reads a bounded local file into a caller-owned output stream. This does not create a main output file; explicitly configured conversion sidecars retain their normal publication policy.</summary>
    public static async Task<OfficeConversionReport> ConvertAsync(string source, Stream destination, OfficeDocumentFormat target,
        OfficeConversionOptions? options, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("A source file is required.", nameof(source));
        string inputPath = Path.GetFullPath(source);
        var settings = (options ?? new()).Snapshot();
        settings.SourceName = ConversionSourceName(inputPath, settings);
        settings.ResourceDirectory = Path.GetDirectoryName(inputPath);
        settings.ResourceSourcePath = inputPath;
        using var input = OfficeConversionInputReader.OpenLocalRead(inputPath);
        var report = await ConvertStreamAsync(input, FormatForExtension(Path.GetExtension(settings.SourceName)), destination, target,
            settings, null, cancellationToken).ConfigureAwait(false);
        return new OfficeConversionReport(inputPath, report.Target, itemsWritten: report.ItemsWritten, elapsed: report.Elapsed, diagnostics: report.Diagnostics);
    }

    /// <summary>Publishes completed bounded bytes under the converter's existing staging and conflict policy.</summary>
    private static async Task<string> PublishConvertedOutputAsync(Stream output, string destination,
        OfficeConversionOptions settings, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        string directory = Path.GetDirectoryName(destination)!;
        await CheckStagingAsync(settings.PublicationGuard, directory, token).ConfigureAwait(false);
        Directory.CreateDirectory(directory);
        await CheckStagingAsync(settings.PublicationGuard, directory, token).ConfigureAwait(false);
        string stage = Path.Combine(directory, "." + Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try {
            using (var stream = new FileStream(stage, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                output.Position = 0;
                await output.CopyToAsync(stream, 81920, token).ConfigureAwait(false);
                stream.Flush(flushToDisk: true);
            }
            return await OfficeFilePublisher.PublishAsync(stage, destination, settings.FileConflictPolicy,
                settings.PublicationGuard, token).ConfigureAwait(false);
        } finally { if (File.Exists(stage)) File.Delete(stage); }
    }
}
