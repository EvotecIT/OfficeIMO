using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf.Filters;

internal static partial class StreamDecoder {
    private static byte[] DecodeFax(PdfDictionary dictionary, int index, byte[] data,
        Dictionary<int, PdfIndirectObject>? objects, int maximumBytes, CancellationToken cancellationToken) {
        PdfDictionary parameters = GetDecodeParms(dictionary, index, objects, cancellationToken) ?? new PdfDictionary();
        ValidateFaxParameters(parameters, objects, cancellationToken);
        int columns = ReadPositiveIntegerParameter(parameters, "Columns", 1728, objects, cancellationToken);
        int width = ReadIntegerParameter(dictionary, "Width", columns, objects, cancellationToken);
        if (width != columns) throw new InvalidDataException("CCITT columns do not match the image width.");
        int rows = ReadIntegerParameter(parameters, "Rows", 0, objects, cancellationToken);
        bool endOfBlock = ReadFaxBoolean(parameters, "EndOfBlock", true, objects, cancellationToken);
        int height = ReadIntegerParameter(dictionary, "Height", 0, objects, cancellationToken);
        // End markers override the filter's Rows hint. Image Height remains the output bound.
        if (endOfBlock && height > 0 || rows == 0) rows = height;
        if (rows <= 0) throw new InvalidDataException("CCITT image decoding requires a row count or image height.");
        long length = (((long)columns + 7) / 8) * rows;
        ThrowIfDecodedLimitExceeded(length, maximumBytes);
        return OfficeFaxDecoder.Decode(data, columns, rows,
            ReadIntegerParameter(parameters, "K", 0, objects, cancellationToken),
            ReadFaxBoolean(parameters, "EndOfLine", false, objects, cancellationToken),
            ReadFaxBoolean(parameters, "EncodedByteAlign", false, objects, cancellationToken),
            ReadFaxBoolean(parameters, "BlackIs1", false, objects, cancellationToken),
            endOfBlock,
            maximumBytes, cancellationToken);
    }

    private static void ValidateFaxParameters(PdfDictionary parameters, Dictionary<int, PdfIndirectObject>? objects, CancellationToken cancellationToken = default) {
        _ = ReadPositiveIntegerParameter(parameters, "Columns", 1728, objects, cancellationToken);
        _ = ReadIntegerParameter(parameters, "K", 0, objects, cancellationToken);
        if (ReadIntegerParameter(parameters, "Rows", 0, objects, cancellationToken) < 0 ||
            ReadIntegerParameter(parameters, "DamagedRowsBeforeError", 0, objects, cancellationToken) < 0) {
            throw new FormatException("Fax row parameters must not be negative.");
        }
        _ = ReadFaxBoolean(parameters, "EndOfLine", false, objects, cancellationToken);
        _ = ReadFaxBoolean(parameters, "EncodedByteAlign", false, objects, cancellationToken);
        _ = ReadFaxBoolean(parameters, "BlackIs1", false, objects, cancellationToken);
        _ = ReadFaxBoolean(parameters, "EndOfBlock", true, objects, cancellationToken);
    }

    private static bool ReadFaxBoolean(PdfDictionary parameters, string name, bool defaultValue,
        Dictionary<int, PdfIndirectObject>? objects, CancellationToken cancellationToken = default) {
        if (!parameters.Items.TryGetValue(name, out PdfObject? value)) return defaultValue;
        PdfObject? resolved = ResolveObject(value, objects, cancellationToken);
        if (resolved is PdfNull) return defaultValue;
        if (resolved is PdfBoolean boolean) return boolean.Value;
        throw new FormatException("Invalid fax boolean parameter: " + name);
    }
}
