using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    /// <summary>Converts a Pages file or directory bundle with cancellation and returns the destination document.</summary>
    public static WordDocument ConvertPagesToWord(string path,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) =>
        ConvertPagesToWordResult(path, readOptions, conversionOptions, cancellationToken).RequireCompleteEditableReconstruction();

    /// <summary>Converts a Pages file or directory bundle with cancellation and retains source evidence.</summary>
    /// <remarks>Cancellation governs loading, semantic projection, and destination construction. Saving is a separate owner operation.</remarks>
    public static PagesToWordResult ConvertPagesToWordResult(string path,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        return ProjectPages(IWorkSourceDocument.Open(path, IWorkDocumentKind.Pages,
            readOptions, cancellationToken), conversionOptions);
    }

    /// <summary>Converts a Pages caller-owned ZIP stream with cancellation and returns the destination document.</summary>
    public static WordDocument ConvertPagesToWord(Stream stream,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) =>
        ConvertPagesToWordResult(stream, readOptions, conversionOptions, cancellationToken).RequireCompleteEditableReconstruction();

    /// <summary>Converts a Pages caller-owned ZIP stream with cancellation and retains source evidence.</summary>
    /// <remarks>Cancellation governs loading, semantic projection, and destination construction. Saving is a separate owner operation.</remarks>
    public static PagesToWordResult ConvertPagesToWordResult(Stream stream,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        return ProjectPages(IWorkSourceDocument.Open(stream, IWorkDocumentKind.Pages,
            readOptions, cancellationToken), conversionOptions);
    }
}
