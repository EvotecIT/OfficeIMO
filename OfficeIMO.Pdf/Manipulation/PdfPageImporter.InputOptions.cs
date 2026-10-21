namespace OfficeIMO.Pdf;

internal static partial class PdfPageImporter {
    /// <summary>Reads bounded remaining streams without closing them and completes page appending with cooperative cancellation.</summary>
    internal static byte[] AppendPages(PdfPageImportOptions options, Stream targetStream, Stream sourceStream,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        var inputs = ReadImportStreams(options, targetStream, sourceStream, targetReadOptions);
        return AppendPages(inputs.Options, inputs.Target.Bytes, inputs.Source.Bytes, inputs.Target.Options, sourcePageNumbers);
    }

    /// <summary>Reads bounded file snapshots and completes page appending with cooperative cancellation.</summary>
    internal static byte[] AppendPages(PdfPageImportOptions options, string targetPath, string sourcePath,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        var inputs = ReadImportFiles(options, targetPath, sourcePath, targetReadOptions);
        return AppendPages(inputs.Options, inputs.Target.Bytes, inputs.Source.Bytes, inputs.Target.Options, sourcePageNumbers);
    }

    /// <summary>Finishes bounded page appending before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void AppendPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, Stream outputStream,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = AppendPages(options, targetPdf, sourcePdf, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options.CancellationToken);
    }

    /// <summary>Finishes bounded page appending before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void AppendPages(PdfPageImportOptions options, Stream targetStream, Stream sourceStream, Stream outputStream,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = AppendPages(options, targetStream, sourceStream, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options.CancellationToken);
    }

    /// <summary>Finishes bounded page appending before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void AppendPages(PdfPageImportOptions options, string targetPath, string sourcePath, Stream outputStream,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = AppendPages(options, targetPath, sourcePath, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options.CancellationToken);
    }

    /// <summary>Finishes bounded page appending before checking cancellation and atomically publishing the destination file.</summary>
    internal static void AppendPages(PdfPageImportOptions options, string targetPath, string sourcePath, string outputPath,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        string destination = ValidateOutputPath(outputPath);
        byte[] output = AppendPages(options, targetPath, sourcePath, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(destination, output, options.CancellationToken);
    }

    /// <summary>Reads bounded remaining streams without closing them and completes page prepending with cooperative cancellation.</summary>
    internal static byte[] PrependPages(PdfPageImportOptions options, Stream targetStream, Stream sourceStream,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        var inputs = ReadImportStreams(options, targetStream, sourceStream, targetReadOptions);
        return PrependPages(inputs.Options, inputs.Target.Bytes, inputs.Source.Bytes, inputs.Target.Options, sourcePageNumbers);
    }

    /// <summary>Reads bounded file snapshots and completes page prepending with cooperative cancellation.</summary>
    internal static byte[] PrependPages(PdfPageImportOptions options, string targetPath, string sourcePath,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        var inputs = ReadImportFiles(options, targetPath, sourcePath, targetReadOptions);
        return PrependPages(inputs.Options, inputs.Target.Bytes, inputs.Source.Bytes, inputs.Target.Options, sourcePageNumbers);
    }

    /// <summary>Finishes bounded page prepending before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void PrependPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, Stream outputStream,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = PrependPages(options, targetPdf, sourcePdf, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options.CancellationToken);
    }

    /// <summary>Finishes bounded page prepending before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void PrependPages(PdfPageImportOptions options, Stream targetStream, Stream sourceStream, Stream outputStream,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = PrependPages(options, targetStream, sourceStream, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options.CancellationToken);
    }

    /// <summary>Finishes bounded page prepending before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void PrependPages(PdfPageImportOptions options, string targetPath, string sourcePath, Stream outputStream,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = PrependPages(options, targetPath, sourcePath, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options.CancellationToken);
    }

    /// <summary>Finishes bounded page prepending before checking cancellation and atomically publishing the destination file.</summary>
    internal static void PrependPages(PdfPageImportOptions options, string targetPath, string sourcePath, string outputPath,
        PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        string destination = ValidateOutputPath(outputPath);
        byte[] output = PrependPages(options, targetPath, sourcePath, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(destination, output, options.CancellationToken);
    }

    /// <summary>Reads bounded remaining streams without closing them and completes page insertion with cooperative cancellation.</summary>
    internal static byte[] InsertPages(PdfPageImportOptions options, Stream targetStream, Stream sourceStream,
        int insertBeforePageNumber, PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        var inputs = ReadImportStreams(options, targetStream, sourceStream, targetReadOptions);
        return InsertPages(inputs.Options, inputs.Target.Bytes, inputs.Source.Bytes, insertBeforePageNumber, inputs.Target.Options, sourcePageNumbers);
    }

    /// <summary>Reads bounded file snapshots and completes page insertion with cooperative cancellation.</summary>
    internal static byte[] InsertPages(PdfPageImportOptions options, string targetPath, string sourcePath,
        int insertBeforePageNumber, PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        var inputs = ReadImportFiles(options, targetPath, sourcePath, targetReadOptions);
        return InsertPages(inputs.Options, inputs.Target.Bytes, inputs.Source.Bytes, insertBeforePageNumber, inputs.Target.Options, sourcePageNumbers);
    }

    /// <summary>Finishes bounded page insertion before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void InsertPages(PdfPageImportOptions options, byte[] targetPdf, byte[] sourcePdf, Stream outputStream,
        int insertBeforePageNumber, PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = InsertPages(options, targetPdf, sourcePdf, insertBeforePageNumber, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options.CancellationToken);
    }

    /// <summary>Finishes bounded page insertion before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void InsertPages(PdfPageImportOptions options, Stream targetStream, Stream sourceStream, Stream outputStream,
        int insertBeforePageNumber, PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = InsertPages(options, targetStream, sourceStream, insertBeforePageNumber, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options.CancellationToken);
    }

    /// <summary>Finishes bounded page insertion before checking cancellation and changing the caller-owned output stream.</summary>
    internal static void InsertPages(PdfPageImportOptions options, string targetPath, string sourcePath, Stream outputStream,
        int insertBeforePageNumber, PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        ValidateWritableOutputStream(outputStream);
        byte[] output = InsertPages(options, targetPath, sourcePath, insertBeforePageNumber, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(outputStream, output, options.CancellationToken);
    }

    /// <summary>Finishes bounded page insertion before checking cancellation and atomically publishing the destination file.</summary>
    internal static void InsertPages(PdfPageImportOptions options, string targetPath, string sourcePath, string outputPath,
        int insertBeforePageNumber, PdfLoadOptions? targetReadOptions, params int[] sourcePageNumbers) {
        string destination = ValidateOutputPath(outputPath);
        byte[] output = InsertPages(options, targetPath, sourcePath, insertBeforePageNumber, targetReadOptions, sourcePageNumbers);
        options.CancellationToken.ThrowIfCancellationRequested();
        WriteOutput(destination, output, options.CancellationToken);
    }

    private static (PdfDocumentSource Target, PdfDocumentSource Source, PdfPageImportOptions Options) ReadImportStreams(
        PdfPageImportOptions options, Stream targetStream, Stream sourceStream, PdfLoadOptions? targetReadOptions) {
        PdfPageImportOptions captured = CaptureImportOptions(options);
        PdfDocumentSource target = PdfDocumentSource.FromRemainingStream(targetStream, targetReadOptions, captured.CancellationToken);
        PdfDocumentSource source = PdfDocumentSource.FromRemainingStream(sourceStream, captured.SourceReadOptions, captured.CancellationToken);
        return (target, source, captured);
    }

    private static (PdfDocumentSource Target, PdfDocumentSource Source, PdfPageImportOptions Options) ReadImportFiles(
        PdfPageImportOptions options, string targetPath, string sourcePath, PdfLoadOptions? targetReadOptions) {
        PdfPageImportOptions captured = CaptureImportOptions(options);
        PdfDocumentSource target = PdfDocumentSource.FromPath(targetPath, targetReadOptions, captured.CancellationToken);
        PdfDocumentSource source = PdfDocumentSource.FromPath(sourcePath, captured.SourceReadOptions, captured.CancellationToken);
        return (target, source, captured);
    }

    private static PdfPageImportOptions CaptureImportOptions(PdfPageImportOptions options) {
        Guard.NotNull(options, nameof(options));
        options.CancellationToken.ThrowIfCancellationRequested();
        return new PdfPageImportOptions {
            CancellationToken = options.CancellationToken,
            FlattenVisualAnnotations = options.FlattenVisualAnnotations,
            SourceReadOptions = PdfLoadOptions.Resolve(options.SourceReadOptions)
        };
    }
}
