using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    /// <summary>Converts a Pages document and returns the Word model together with the typed source projection and loss report.</summary>
    /// <remarks>The result keeps the opened source package and its typed projection, which the common report does not carry.
    /// Use <see cref="PagesToWordResult.RequireCompleteEditableReconstruction"/> to reject visual fallbacks and incomplete reconstruction.</remarks>
    public static PagesToWordResult ToWordResult(this IWorkSourceDocument package, IWorkConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (package == null) throw new ArgumentNullException(nameof(package));
        cancellationToken.ThrowIfCancellationRequested();
        return ProjectPages(package, options);
    }

    /// <summary>Converts a Pages document into the editable Word model.</summary>
    public static WordDocument ToWord(this IWorkSourceDocument package, PagesWordExportOptions? options = null, CancellationToken cancellationToken = default) =>
        package.ToWord(options, out _, cancellationToken);

    /// <summary>Converts a Pages document into the editable Word model and returns the completed conversion report.</summary>
    public static WordDocument ToWord(this IWorkSourceDocument package, PagesWordExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (package == null) throw new ArgumentNullException(nameof(package));
        cancellationToken.ThrowIfCancellationRequested();
        PagesToWordResult result = package.ToWordResult(options?.ConversionOptions, cancellationToken);
        WordDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a Pages document into the editable Word model, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<WordDocument>> ToWordAsync(this IWorkSourceDocument package, PagesWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (package.ToWord(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsWord(this IWorkSourceDocument package, string path, PagesWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = package.ToWord(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsWord(this IWorkSourceDocument package, Stream stream, PagesWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = package.ToWord(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this IWorkSourceDocument package, string path, PagesWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using WordDocument value = package.ToWord(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this IWorkSourceDocument package, Stream stream, PagesWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using WordDocument value = package.ToWord(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
