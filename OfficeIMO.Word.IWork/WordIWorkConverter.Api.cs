using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    /// <summary>Opens and converts a Pages file into the normal editable Word model.</summary>
    public static WordDocument ConvertPagesToWord(string path,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) =>
        ConvertPagesToWordResult(path, readOptions, conversionOptions).Value;

    /// <summary>Opens and converts a Pages stream into the normal editable Word model.</summary>
    public static WordDocument ConvertPagesToWord(Stream stream,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) =>
        ConvertPagesToWordResult(stream, readOptions, conversionOptions).Value;

    /// <summary>Opens and converts a Pages file, returning the Word document with source evidence.</summary>
    public static PagesToWordResult ConvertPagesToWordResult(string path,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return ProjectPages(
            IWorkSourceDocument.Open(path, IWorkDocumentKind.Pages, readOptions), conversionOptions);
    }

    /// <summary>Opens and converts a Pages stream, returning the Word document with source evidence.</summary>
    public static PagesToWordResult ConvertPagesToWordResult(Stream stream,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return ProjectPages(
            IWorkSourceDocument.Open(stream, IWorkDocumentKind.Pages, readOptions), conversionOptions);
    }

    /// <summary>Converts an opened Pages source into the normal editable Word model.</summary>
    public static WordDocument ToWordDocument(this IWorkSourceDocument source,
        IWorkConversionOptions? options = null) =>
        ToWordDocumentResult(source, options).Value;

    /// <summary>Converts an opened Pages source and returns the Word document with its projection and loss report.</summary>
    public static PagesToWordResult ToWordDocumentResult(this IWorkSourceDocument source,
        IWorkConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return ProjectPages(source, options);
    }
}
