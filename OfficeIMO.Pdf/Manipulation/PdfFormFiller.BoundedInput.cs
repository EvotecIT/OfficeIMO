namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    /// <summary>Fills fields from a bounded snapshot at the caller's current stream position.</summary>
    internal static byte[] FillFields(Stream stream, IReadOnlyDictionary<string, string> fieldValues,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FillFields(ReadStream(stream, nameof(stream), options, readOptions), fieldValues, options, readOptions);

    /// <summary>Fills typed fields from a bounded snapshot at the caller's current stream position.</summary>
    internal static byte[] FillFields(Stream stream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FillFields(ReadStream(stream, nameof(stream), options, readOptions), fieldValues, options, readOptions);

    /// <summary>Flattens fields from a bounded snapshot without owning the caller's stream.</summary>
    internal static byte[] FlattenFields(Stream stream, PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FlattenFields(ReadStream(stream, nameof(stream), options, readOptions), options, readOptions);

    /// <summary>Fills and flattens fields from one bounded input snapshot.</summary>
    internal static byte[] FillAndFlattenFields(Stream stream, IReadOnlyDictionary<string, string> fieldValues,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FillAndFlattenFields(ReadStream(stream, nameof(stream), options, readOptions), fieldValues, options, readOptions);

    /// <summary>Fills and flattens typed fields from one bounded input snapshot.</summary>
    internal static byte[] FillAndFlattenFields(Stream stream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FillAndFlattenFields(ReadStream(stream, nameof(stream), options, readOptions), fieldValues, options, readOptions);

    /// <summary>Fills fields from a file snapshot admitted against the supplied read limits.</summary>
    internal static byte[] FillFieldsToBytes(string path, IReadOnlyDictionary<string, string> fieldValues,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FillFields(ReadPath(path, options, readOptions), fieldValues, options, readOptions);

    /// <summary>Flattens fields from a file snapshot admitted against the supplied read limits.</summary>
    internal static byte[] FlattenFieldsToBytes(string path, PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FlattenFields(ReadPath(path, options, readOptions), options, readOptions);

    /// <summary>Fills and flattens fields from a file snapshot admitted against the supplied read limits.</summary>
    internal static byte[] FillAndFlattenFieldsToBytes(string path, IReadOnlyDictionary<string, string> fieldValues,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FillAndFlattenFields(ReadPath(path, options, readOptions), fieldValues, options, readOptions);
}
