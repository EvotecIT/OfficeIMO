using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

public static partial class PowerPointIWorkConverter {
    /// <summary>Converts a Keynote file or directory bundle with cancellation and returns the destination document.</summary>
    public static PowerPointPresentation ConvertKeynoteToPowerPoint(string path,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) =>
        ConvertKeynoteToPowerPointResult(path, readOptions, conversionOptions, cancellationToken).Value;

    /// <summary>Converts a Keynote file or directory bundle with cancellation and retains source evidence.</summary>
    /// <remarks>Cancellation governs loading, semantic projection, and destination construction. Saving is a separate owner operation.</remarks>
    public static KeynoteToPowerPointResult ConvertKeynoteToPowerPointResult(string path,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        return ProjectKeynote(IWorkSourceDocument.Open(path, IWorkDocumentKind.Keynote,
            readOptions, cancellationToken), conversionOptions);
    }

    /// <summary>Converts a Keynote caller-owned ZIP stream with cancellation and returns the destination document.</summary>
    public static PowerPointPresentation ConvertKeynoteToPowerPoint(Stream stream,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) =>
        ConvertKeynoteToPowerPointResult(stream, readOptions, conversionOptions, cancellationToken).Value;

    /// <summary>Converts a Keynote caller-owned ZIP stream with cancellation and retains source evidence.</summary>
    /// <remarks>Cancellation governs loading, semantic projection, and destination construction. Saving is a separate owner operation.</remarks>
    public static KeynoteToPowerPointResult ConvertKeynoteToPowerPointResult(Stream stream,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        return ProjectKeynote(IWorkSourceDocument.Open(stream, IWorkDocumentKind.Keynote,
            readOptions, cancellationToken), conversionOptions);
    }
}
