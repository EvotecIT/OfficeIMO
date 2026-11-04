using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal;

/// <summary>Completes a converted value and writes converted output with the common report contract.</summary>
internal static class OfficeConversionOutput {
    private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

    /// <summary>Applies the loss policy and returns the value; a rejected disposable value is disposed.</summary>
    internal static T Complete<T>(T value, OfficeConversionReport source, bool failOnLoss, out OfficeConversionReport report) {
        report = source;
        if (failOnLoss) {
            try { source.ThrowIfLoss(); }
            catch { (value as IDisposable)?.Dispose(); throw; }
        }
        return value;
    }

    internal static byte[] Text(string text, Encoding? encoding = null) => (encoding ?? Utf8NoBom).GetBytes(text);

    internal static OfficeConversionReport Save(OfficeConversionReport report, string path, byte[] bytes) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Output path cannot be empty.", nameof(path));
        string fullPath = Path.GetFullPath(path);
        OfficeConversionFailures.Run(report, () => OfficeFileCommit.WriteAllBytes(fullPath, bytes));
        return WithPath(report, fullPath);
    }

    internal static OfficeConversionReport Save(OfficeConversionReport report, Stream stream, byte[] bytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        OfficeConversionFailures.Run(report, () => OfficeStreamWriter.WriteAllBytes(stream, bytes));
        return report;
    }

    internal static async Task<OfficeConversionReport> SaveAsync(OfficeConversionReport report, string path, byte[] bytes,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Output path cannot be empty.", nameof(path));
        string fullPath = Path.GetFullPath(path);
        await OfficeConversionFailures.RunAsync(report, () => OfficeFileCommit.WriteAllBytesAsync(fullPath, bytes,
            cancellationToken: cancellationToken)).ConfigureAwait(false);
        return WithPath(report, fullPath);
    }

    internal static async Task<OfficeConversionReport> SaveAsync(OfficeConversionReport report, Stream stream, byte[] bytes,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        await OfficeConversionFailures.RunAsync(report, () => OfficeStreamWriter.WriteAllBytesAsync(stream, bytes,
            cancellationToken)).ConfigureAwait(false);
        return report;
    }

    /// <summary>Runs CPU-bound conversion work on the thread pool and pairs the value with its report.</summary>
    internal static Task<OfficeConversionResult<T>> RunAsync<T>(Func<(T Value, OfficeConversionReport Report)> convert,
        CancellationToken cancellationToken) =>
        Task.Run(() => {
            var converted = convert();
            return new OfficeConversionResult<T>(converted.Value, converted.Report);
        }, cancellationToken);

    private static OfficeConversionReport WithPath(OfficeConversionReport report, string path) =>
        new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
}
