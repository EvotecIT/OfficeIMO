using System.Globalization;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

/// <summary>Reads bounded interchange records without splitting or retaining the entire decoded source.</summary>
internal sealed class TextSpreadsheetReader : IDisposable {
    private readonly StreamReader _reader;
    private readonly CancellationToken _token;
    private readonly int _lineLimit;
    private int _lines;
    private int _textCharacters;

    internal TextSpreadsheetReader(byte[] data, LegacySpreadsheetImportOptions options, CancellationToken token) {
        Limits = options.Limits;
        _token = token;
        _lineLimit = (int)Math.Min(Limits.MaxInputBytes, Math.Max(33_791L, (long)Limits.MaxTextCharacters + 1024));
        Encoding encoding = ResolveEncoding(data, options.TextEncoding, out int offset);
        _reader = new StreamReader(new MemoryStream(data, offset, data.Length - offset, writable: false), encoding, detectEncodingFromByteOrderMarks: false);
    }

    internal OfficeLegacyImportLimits Limits { get; }

    internal string? ReadLine() {
        _token.ThrowIfCancellationRequested();
        if (_reader.Peek() < 0) return null;
        if (++_lines > Limits.MaxRecords) throw new InvalidDataException("Text spreadsheet exceeds the configured record limit.");
        var line = new StringBuilder();
        int next;
        while ((next = _reader.Read()) >= 0) {
            if ((line.Length & 1023) == 0) _token.ThrowIfCancellationRequested();
            if (next == '\n') break;
            if (next == '\r') {
                if (_reader.Peek() == '\n') _reader.Read();
                break;
            }
            if (next == 0) throw new InvalidDataException("Text spreadsheet contains a NUL character.");
            if (line.Length >= _lineLimit) throw new InvalidDataException("Text spreadsheet record exceeds the configured text limit.");
            line.Append((char)next);
        }
        return line.ToString();
    }

    internal string RequiredLine() => ReadLine() ?? throw new InvalidDataException("Text spreadsheet ended inside a record.");

    internal void RequireEnd() {
        string? line;
        while ((line = ReadLine()) != null) {
            if (!string.IsNullOrWhiteSpace(line)) throw new InvalidDataException("Text spreadsheet contains data after its terminator.");
        }
    }

    internal void AddCell(LegacySpreadsheetModel model, LegacySpreadsheetSheet sheet, int row, int column, object? value) {
        _token.ThrowIfCancellationRequested();
        if (row < 1 || row > 1_048_576 || column < 1 || column > 16_384) throw new InvalidDataException("Text spreadsheet cell address is outside the Excel worksheet bounds.");
        if (model.RecoveredCellCount >= Limits.MaxItems) throw new InvalidDataException("Text spreadsheet exceeds the configured cell limit.");
        if (value is string text) {
            if (text.Length > 32_767) throw new InvalidDataException("Text spreadsheet cell exceeds Excel's 32,767 character limit.");
            RetainText(text);
        }
        sheet.Cells.Add(new LegacySpreadsheetCell(row, column, value));
        model.RecoveredCellCount++;
    }

    internal void RetainText(string text) {
        if (text.Length > Limits.MaxTextCharacters - _textCharacters) throw new InvalidDataException("Text spreadsheet exceeds the configured text limit.");
        _textCharacters += text.Length;
    }

    internal static int Coordinate(string value, int maximum) {
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int coordinate) || coordinate < 1 || coordinate > maximum)
            throw new InvalidDataException("Invalid text spreadsheet cell address.");
        return coordinate;
    }

    internal static double Number(string value) {
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) || double.IsNaN(number) || double.IsInfinity(number))
            throw new InvalidDataException("Text spreadsheet contains an invalid or non-finite numeric value.");
        return number;
    }

    internal static bool HasHeader(byte[] data, string header) {
        try {
            Encoding encoding = ResolveEncoding(data, null, out int offset);
            int byteCount = encoding.GetByteCount(header);
            return data.Length - offset >= byteCount && encoding.GetString(data, offset, byteCount) == header;
        } catch (DecoderFallbackException) {
            return false;
        }
    }

    private static Encoding ResolveEncoding(byte[] data, Encoding? configured, out int offset) {
        offset = 0;
        if (OfficeLegacyImportBuffer.StartsWith(data, 0xEF, 0xBB, 0xBF)) { offset = 3; return new UTF8Encoding(false, true); }
        if (OfficeLegacyImportBuffer.StartsWith(data, 0xFF, 0xFE, 0, 0)) { offset = 4; return new UTF32Encoding(false, false, true); }
        if (OfficeLegacyImportBuffer.StartsWith(data, 0, 0, 0xFE, 0xFF)) { offset = 4; return new UTF32Encoding(true, false, true); }
        if (OfficeLegacyImportBuffer.StartsWith(data, 0xFF, 0xFE)) { offset = 2; return new UnicodeEncoding(false, false, true); }
        if (OfficeLegacyImportBuffer.StartsWith(data, 0xFE, 0xFF)) { offset = 2; return new UnicodeEncoding(true, false, true); }
        var encoding = (Encoding)(configured ?? CodePagesEncodingProvider.Instance.GetEncoding(1252)!).Clone();
        encoding.DecoderFallback = DecoderFallback.ExceptionFallback;
        return encoding;
    }

    public void Dispose() => _reader.Dispose();
}

internal abstract class TextSpreadsheetAdapterBase : LegacySpreadsheetAdapterBase {
    public sealed override LegacySpreadsheetModel Parse(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) =>
        Parse(data, new LegacySpreadsheetImportOptions { Limits = limits }, cancellationToken);

    public sealed override LegacySpreadsheetModel Parse(byte[] data, LegacySpreadsheetImportOptions options, CancellationToken cancellationToken) {
        using var reader = new TextSpreadsheetReader(data, options, cancellationToken);
        return Parse(reader);
    }

    protected abstract LegacySpreadsheetModel Parse(TextSpreadsheetReader reader);
}
