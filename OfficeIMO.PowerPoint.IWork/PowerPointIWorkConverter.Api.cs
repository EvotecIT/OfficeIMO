using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

public static partial class PowerPointIWorkConverter {
    /// <summary>Converts a Keynote presentation and returns the PowerPoint model together with the typed source projection and loss report.</summary>
    /// <remarks>The result keeps the opened source package and its typed projection, which the common report does not carry.
    /// Use <see cref="KeynoteToPowerPointResult.RequireCompleteEditableReconstruction"/> to reject visual fallbacks and incomplete reconstruction.</remarks>
    public static KeynoteToPowerPointResult ToPowerPointResult(this IWorkSourceDocument package, IWorkConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(package);
        cancellationToken.ThrowIfCancellationRequested();
        return ProjectKeynote(package, options);
    }

    /// <summary>Converts a Keynote presentation into the editable PowerPoint model.</summary>
    public static PowerPointDocument ToPowerPoint(this IWorkSourceDocument package, KeynotePowerPointExportOptions? options = null, CancellationToken cancellationToken = default) =>
        package.ToPowerPoint(options, out _, cancellationToken);

    /// <summary>Converts a Keynote presentation into the editable PowerPoint model and returns the completed conversion report.</summary>
    public static PowerPointDocument ToPowerPoint(this IWorkSourceDocument package, KeynotePowerPointExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(package);
        cancellationToken.ThrowIfCancellationRequested();
        KeynoteToPowerPointResult result = package.ToPowerPointResult(options?.ConversionOptions, cancellationToken);
        PowerPointDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a Keynote presentation into the editable PowerPoint model, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<PowerPointDocument>> ToPowerPointAsync(this IWorkSourceDocument package, KeynotePowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (package.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this IWorkSourceDocument package, string path, KeynotePowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PowerPointDocument value = package.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this IWorkSourceDocument package, Stream stream, KeynotePowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PowerPointDocument value = package.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this IWorkSourceDocument package, string path, KeynotePowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using PowerPointDocument value = package.ToPowerPoint(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this IWorkSourceDocument package, Stream stream, KeynotePowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using PowerPointDocument value = package.ToPowerPoint(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
