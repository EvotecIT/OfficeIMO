using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf.Filters;

internal static partial class StreamDecoder {
    private static byte[] DecodeFax(PdfDictionary dictionary, int index, byte[] data,
        Dictionary<int, PdfIndirectObject>? objects, int maximumBytes, CancellationToken cancellationToken) {
        PdfDictionary parameters = GetDecodeParms(dictionary, index, objects) ?? new PdfDictionary();
        ValidateFaxParameters(parameters, objects);
        int columns = ReadPositiveIntegerParameter(parameters, "Columns", 1728, objects);
        int width = ReadIntegerParameter(dictionary, "Width", columns, objects);
        if (width != columns) throw new InvalidDataException("CCITT columns do not match the image width.");
        int rows = ReadIntegerParameter(parameters, "Rows", 0, objects);
        if (rows == 0) rows = ReadIntegerParameter(dictionary, "Height", 0, objects);
        if (rows <= 0) throw new InvalidDataException("CCITT image decoding requires a row count or image height.");
        long length = (((long)columns + 7) / 8) * rows;
        ThrowIfDecodedLimitExceeded(length, maximumBytes);
        return OfficeFaxDecoder.Decode(data, columns, rows,
            ReadIntegerParameter(parameters, "K", 0, objects),
            ReadFaxBoolean(parameters, "EndOfLine", false, objects),
            ReadFaxBoolean(parameters, "EncodedByteAlign", false, objects),
            ReadFaxBoolean(parameters, "BlackIs1", false, objects),
            ReadFaxBoolean(parameters, "EndOfBlock", true, objects),
            maximumBytes, cancellationToken);
    }

    private static void ValidateFaxParameters(PdfDictionary parameters, Dictionary<int, PdfIndirectObject>? objects) {
        _ = ReadPositiveIntegerParameter(parameters, "Columns", 1728, objects);
        _ = ReadIntegerParameter(parameters, "K", 0, objects);
        if (ReadIntegerParameter(parameters, "Rows", 0, objects) < 0 ||
            ReadIntegerParameter(parameters, "DamagedRowsBeforeError", 0, objects) < 0) {
            throw new FormatException("Fax row parameters must not be negative.");
        }
        _ = ReadFaxBoolean(parameters, "EndOfLine", false, objects);
        _ = ReadFaxBoolean(parameters, "EncodedByteAlign", false, objects);
        _ = ReadFaxBoolean(parameters, "BlackIs1", false, objects);
        _ = ReadFaxBoolean(parameters, "EndOfBlock", true, objects);
    }

    private static bool ReadFaxBoolean(PdfDictionary parameters, string name, bool defaultValue,
        Dictionary<int, PdfIndirectObject>? objects) {
        if (!parameters.Items.TryGetValue(name, out PdfObject? value)) return defaultValue;
        PdfObject? resolved = ResolveObject(value, objects);
        if (resolved is PdfNull) return defaultValue;
        if (resolved is PdfBoolean boolean) return boolean.Value;
        throw new FormatException("Invalid fax boolean parameter: " + name);
    }
}
