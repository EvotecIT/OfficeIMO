using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    /// <summary>Converts a presentation to an in-memory ODP presentation; the LossPolicy option rejects loss.</summary>
    public static OdpPresentation ToOpenDocument(this PowerPointDocument source, PowerPointOpenDocumentConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        source.ToOpenDocument(options, out _, cancellationToken);

    /// <summary>Converts a presentation to an in-memory ODP presentation; the LossPolicy option rejects loss and returns the completed conversion report.</summary>
    public static OdpPresentation ToOpenDocument(this PowerPointDocument source, PowerPointOpenDocumentConversionOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence(options);
        OdpPresentation value = result.Value;
        OfficeConversionReport evidence = result.Report;
        return OfficeConversionOutput.Complete(value, evidence, false, out report);
    }

    /// <summary>Converts a presentation to an in-memory ODP presentation; the LossPolicy option rejects loss, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<OdpPresentation>> ToOpenDocumentAsync(this PowerPointDocument source, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (source.ToOpenDocument(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsOpenDocument(this PowerPointDocument source, string path, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using OdpPresentation value = source.ToOpenDocument(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsOpenDocument(this PowerPointDocument source, Stream stream, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using OdpPresentation value = source.ToOpenDocument(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOpenDocumentAsync(this PowerPointDocument source, string path, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using OdpPresentation value = source.ToOpenDocument(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOpenDocumentAsync(this PowerPointDocument source, Stream stream, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using OdpPresentation value = source.ToOpenDocument(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts an ODP presentation to an in-memory presentation; the LossPolicy option rejects loss.</summary>
    public static PowerPointDocument ToPowerPoint(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        source.ToPowerPoint(options, out _, cancellationToken);

    /// <summary>Converts an ODP presentation to an in-memory presentation; the LossPolicy option rejects loss and returns the completed conversion report.</summary>
    public static PowerPointDocument ToPowerPoint(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        OdfConversionResult<PowerPointDocument> result = source.ConvertToPowerPointWithEvidence(options);
        PowerPointDocument value = result.Value;
        OfficeConversionReport evidence = result.Report;
        return OfficeConversionOutput.Complete(value, evidence, false, out report);
    }

    /// <summary>Converts an ODP presentation to an in-memory presentation; the LossPolicy option rejects loss, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<PowerPointDocument>> ToPowerPointAsync(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (source.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PowerPointDocument value = source.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PowerPointDocument value = source.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using PowerPointDocument value = source.ToPowerPoint(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using PowerPointDocument value = source.ToPowerPoint(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
