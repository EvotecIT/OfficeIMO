using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

public static partial class PowerPointIWorkConverter {
    /// <summary>Opens and converts a Keynote file into the normal editable PowerPoint model.</summary>
    public static PowerPointPresentation ConvertKeynoteToPowerPoint(string path,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) =>
        ConvertKeynoteToPowerPointResult(path, readOptions, conversionOptions).Value;

    /// <summary>Opens and converts a Keynote stream into the normal editable PowerPoint model.</summary>
    public static PowerPointPresentation ConvertKeynoteToPowerPoint(Stream stream,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) =>
        ConvertKeynoteToPowerPointResult(stream, readOptions, conversionOptions).Value;

    /// <summary>Opens and converts a Keynote file, returning the presentation with source evidence.</summary>
    public static KeynoteToPowerPointResult ConvertKeynoteToPowerPointResult(string path,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return ProjectKeynote(
            IWorkSourceDocument.Open(path, IWorkDocumentKind.Keynote, readOptions), conversionOptions);
    }

    /// <summary>Opens and converts a Keynote stream, returning the presentation with source evidence.</summary>
    public static KeynoteToPowerPointResult ConvertKeynoteToPowerPointResult(Stream stream,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return ProjectKeynote(
            IWorkSourceDocument.Open(stream, IWorkDocumentKind.Keynote, readOptions), conversionOptions);
    }

    /// <summary>Converts an opened Keynote source into the normal editable PowerPoint model.</summary>
    public static PowerPointPresentation ToPowerPointPresentation(this IWorkSourceDocument source,
        IWorkConversionOptions? options = null) =>
        ToPowerPointPresentationResult(source, options).Value;

    /// <summary>Converts an opened Keynote source and returns the presentation with its projection and loss report.</summary>
    public static KeynoteToPowerPointResult ToPowerPointPresentationResult(
        this IWorkSourceDocument source, IWorkConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return ProjectKeynote(source, options);
    }
}
