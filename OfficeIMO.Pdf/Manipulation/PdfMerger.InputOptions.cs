using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfMerger {
    /// <summary>Merges bounded remaining stream snapshots, retaining caller ownership and observing cancellation through serialization.</summary>
    internal static byte[] Merge(IEnumerable<Stream> streams, PdfLoadOptions? readOptions, CancellationToken cancellationToken) =>
        MergeStreams(streams, options: null, readOptions, cancellationToken);

    /// <summary>Merges bounded remaining stream snapshots with explicit document-structure policy and cancellation.</summary>
    internal static byte[] Merge(PdfMergeOptions options, IEnumerable<Stream> streams, PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        Guard.NotNull(options, nameof(options));
        return MergeStreams(streams, options, readOptions, cancellationToken);
    }

    /// <summary>Completes source admission and merging before checking cancellation and changing the caller's output.</summary>
    internal static void Merge(IEnumerable<Stream> streams, Stream outputStream, PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        byte[] output = Merge(streams, readOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeStreamWriter.CopyBytes(outputStream, output, cancellationToken);
    }

    private static byte[] MergeStreams(IEnumerable<Stream> streams, PdfMergeOptions? options, PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        Guard.NotNull(streams, nameof(streams));
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions capturedOptions = PdfLoadOptions.Resolve(readOptions);
        var inputs = new List<Stream>();
        foreach (Stream stream in streams) {
            cancellationToken.ThrowIfCancellationRequested();
            if (stream is null) throw new ArgumentException("PDF stream input " + inputs.Count + " cannot be null.", nameof(streams));
            if (!stream.CanRead) throw new ArgumentException("PDF stream input " + inputs.Count + " must be readable.", nameof(streams));
            inputs.Add(stream);
        }
        if (inputs.Count == 0) throw new ArgumentException("At least one PDF stream must be supplied.", nameof(streams));
        var sources = new List<byte[]>(inputs.Count);
        foreach (Stream stream in inputs) {
            cancellationToken.ThrowIfCancellationRequested();
            sources.Add(PdfDocumentSource.FromRemainingStream(stream, capturedOptions, cancellationToken).Bytes);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return MergeCore(sources, primarySourceIndex: 0, options,
            Enumerable.Repeat(capturedOptions, sources.Count).ToArray(), cancellationToken: cancellationToken).OwnedBytes;
    }

    /// <summary>Merges bounded file snapshots with cancellation through admission, preparation and serialization.</summary>
    internal static byte[] MergeFilesToBytes(IEnumerable<string> inputPaths, PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) =>
        MergeFileSnapshots(inputPaths, options: null, readOptions, cancellationToken);

    /// <summary>Merges bounded file snapshots using caller-selected policy, read limits and cancellation.</summary>
    internal static byte[] MergeFilesToBytes(PdfMergeOptions options, IEnumerable<string> inputPaths, PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        Guard.NotNull(options, nameof(options));
        return MergeFileSnapshots(inputPaths, options, readOptions, cancellationToken);
    }

    /// <summary>Completes every input and output before checking cancellation and publishing the destination file.</summary>
    internal static void MergeFiles(IEnumerable<string> inputPaths, string outputPath, PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        string destination = ValidateOutputPath(outputPath);
        byte[] output = MergeFilesToBytes(inputPaths, readOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        string? directory = Path.GetDirectoryName(destination);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.Write(destination, stream => OfficeStreamWriter.CopyBytes(stream, output, cancellationToken), cancellationToken);
    }

    private static byte[] MergeFileSnapshots(IEnumerable<string> inputPaths, PdfMergeOptions? options, PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        Guard.NotNull(inputPaths, nameof(inputPaths));
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions capturedOptions = PdfLoadOptions.Resolve(readOptions);
        var sources = new List<byte[]>();
        foreach (string path in inputPaths) {
            cancellationToken.ThrowIfCancellationRequested();
            sources.Add(PdfDocumentSource.FromPath(path, capturedOptions, cancellationToken).Bytes);
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (sources.Count == 0) throw new ArgumentException("At least one PDF input path must be supplied.", nameof(inputPaths));
        return MergeCore(sources, primarySourceIndex: 0, options,
            Enumerable.Repeat(capturedOptions, sources.Count).ToArray(), cancellationToken: cancellationToken).OwnedBytes;
    }
}
