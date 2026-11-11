#if NET8_0_OR_GREATER
using OfficeIMO.CSV;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// CSV and TSV output from a JSON array of rows. Pass two walks <see cref="Utf8JsonReader"/> tokens straight into the
/// column trie and hands each row's UTF-8 cells to the CSV writer, without the event layer, so a row costs about what
/// parsing it does. It follows <see cref="Walk"/> exactly: the same cells, diagnostics, locations and errors.
/// </summary>
public sealed partial class StructuredDataReader {
    private const int JsonBufferSize = 64 * 1024;
    private Stream? _jsonInput;
    private long _jsonInputStart;
    private int _visit;
    private byte[] _nameScratch = Array.Empty<byte>();
    private byte[] _valueScratch = Array.Empty<byte>();
    private long _tokenBase;
    private readonly List<(string Code, OfficeConversionLossKind Loss, string Message, string? Location)> _rowReports = new();

    /// <summary>Whether <see cref="WriteJsonRows"/> serves this reader: one JSON document whose root array holds the rows.</summary>
    internal bool CanWriteJsonRows =>
        _utf8 && _jsonInput != null && _mode == Mode.Rows && _select == null && !_multiDocument && _tables.Count == 1
        && _context.Options.ArrayColumns != DataArrayColumns.Join
        && (_context.Options.JsonReader?.CommentHandling ?? JsonCommentHandling.Disallow) != JsonCommentHandling.Allow;

    /// <summary>Writes every row to <paramref name="writer"/>, whose columns are already established.</summary>
    internal void WriteJsonRows(CsvRowWriter writer) {
        Stream input = _jsonInput!;
        input.Position = _jsonInputStart;
        int count = FieldCount;
        JsonReaderOptions options = _context.Options.JsonReader ?? new JsonReaderOptions { MaxDepth = 256 };
        var state = new JsonReaderState(options);
        byte[] buffer = ArrayPool<byte>.Shared.Rent(JsonBufferSize);
        int start = 0, end = 0;
        long offset = 0;
        bool final = false, bomChecked = false, rootOpen = false;
        try {
            while (true) {
                if (!final) Fill(input, ref buffer, ref start, ref end, ref offset, ref final);
                if (!bomChecked && (end - start >= 3 || final)) {
                    bomChecked = true;
                    if (offset == 0 && end >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF) start = 3;
                }

                if (!bomChecked) continue;
                _tokenBase = offset + start;
                var reader = new Utf8JsonReader(new ReadOnlySpan<byte>(buffer, start, end - start), final, state);
                bool more = false;
                try {
                    while (true) {
                        JsonReaderState rowState = reader.CurrentState;
                        int rowStart = (int)reader.BytesConsumed;
                        if (!reader.Read()) {
                            state = reader.CurrentState;
                            start += (int)reader.BytesConsumed;
                            more = !final;
                            break;
                        }

                        if (!rootOpen) {
                            rootOpen = true;
                            continue;
                        }

                        if (reader.CurrentDepth == 0) continue;
                        _context.CancellationToken.ThrowIfCancellationRequested();
                        BeginRow();
                        if (!WalkJson(ref reader, Current.Root)) {
                            state = rowState;
                            start += rowStart;
                            more = true;
                            break;
                        }

                        EndRow();
                        writer.WriteUtf8Row(count, this, static (rows, ordinal) => rows.GetUtf8(ordinal));
                    }
                } catch (JsonException error) {
                    throw new InvalidDataException("The JSON input is invalid near byte "
                        + (_tokenBase + reader.BytesConsumed).ToString(CultureInfo.InvariantCulture) + ": " + error.Message, error);
                }

                if (!more) return;
            }
        } finally {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    // Keeps the unread bytes, which start at a row boundary, and reads more after them.
    private static void Fill(Stream input, ref byte[] buffer, ref int start, ref int end, ref long offset, ref bool final) {
        int remaining = end - start;
        if (start > 0) {
            Buffer.BlockCopy(buffer, start, buffer, 0, remaining);
            offset += start;
            start = 0;
            end = remaining;
        }

        if (end == buffer.Length) {
            byte[] larger = ArrayPool<byte>.Shared.Rent(buffer.Length * 2);
            Buffer.BlockCopy(buffer, 0, larger, 0, end);
            ArrayPool<byte>.Shared.Return(buffer);
            buffer = larger;
        }

        int read = input.Read(buffer, end, buffer.Length - end);
        if (read == 0) final = true;
        else end += read;
    }

    private void BeginRow() {
        _cellLengths.AsSpan().Fill(-1);
        _cellBytesLength = 0;
        Array.Clear(_kinds, 0, _kinds.Length);
        _rowReports.Clear();
    }

    // A row's findings are kept until the row is complete, because a row cut by the end of the buffer is read again.
    private void EndRow() {
        foreach ((string code, OfficeConversionLossKind loss, string message, string? location) in _rowReports) _context.Report(code, loss, message, location);
        _row++;
        _context.ItemsWritten++;
    }

    private string Location(long token) => "byte " + (_tokenBase + token).ToString(CultureInfo.InvariantCulture);

    /// <summary>Walks the value at the current token into cells; false when the buffer ends inside it.</summary>
    private bool WalkJson(ref Utf8JsonReader reader, TableColumnNode? node) {
        switch (reader.TokenType) {
            case JsonTokenType.StartObject: {
                int visit = ++_visit;
                List<string>? unknown = null;
                while (true) {
                    if (!reader.Read()) return false;
                    if (reader.TokenType == JsonTokenType.EndObject) return true;
                    ReadOnlySpan<byte> name = Unescaped(ref reader, names: true);
                    TableColumnNode? child = node?.FindName(name);
                    if (child != null) {
                        if (child.Visit == visit) throw Duplicate(name, reader.TokenStartIndex);
                        child.Visit = visit;
                    } else {
                        string key = Encoding.UTF8.GetString(name);
                        unknown ??= new List<string>();
                        if (unknown.Contains(key)) throw Duplicate(name, reader.TokenStartIndex);
                        unknown.Add(key);
                        if (node != null && _sampled) _latePath = (node.Path.Length == 0 ? string.Empty : node.Path + ".") + key;
                    }

                    if (!reader.Read() || !WalkJson(ref reader, child)) return false;
                }
            }
            case JsonTokenType.StartArray: {
                int index = 0;
                while (true) {
                    if (!reader.Read()) return false;
                    if (reader.TokenType == JsonTokenType.EndArray) return true;
                    if (!WalkJson(ref reader, node?.FindIndex(index))) return false;
                    index++;
                }
            }
            case JsonTokenType.String:
                Scalar(reader.TokenStartIndex, node, DataScalarKind.String, Unescaped(ref reader, names: false));
                return true;
            case JsonTokenType.Number: {
                ReadOnlySpan<byte> raw = reader.ValueSpan;
                Scalar(reader.TokenStartIndex, node, raw.IndexOfAny((byte)'.', (byte)'e', (byte)'E') < 0 ? DataScalarKind.Integer : DataScalarKind.Float, raw);
                return true;
            }
            case JsonTokenType.True:
                Scalar(reader.TokenStartIndex, node, DataScalarKind.Boolean, "true"u8);
                return true;
            case JsonTokenType.False:
                Scalar(reader.TokenStartIndex, node, DataScalarKind.Boolean, "false"u8);
                return true;
            default:
                Scalar(reader.TokenStartIndex, node, DataScalarKind.Null, default);
                return true;
        }
    }

    private void Scalar(long token, TableColumnNode? node, DataScalarKind kind, ReadOnlySpan<byte> utf8) {
        if (kind == DataScalarKind.String && !System.Text.Unicode.Utf8.IsValid(utf8)) {
            throw new InvalidDataException("The JSON input has invalid text near " + Location(token) + ": a string is not valid UTF-8.");
        }

        if (node is not { Column: >= 0 }) {
            if (_sampled && kind != DataScalarKind.Null) {
                string? path = node?.Path is { Length: > 0 } known ? known : _latePath;
                throw LateColumnError(path, Location(token));
            }

            return;
        }

        int column = node.Column;
        if (kind == DataScalarKind.Null && node.Kind is null or DataScalarKind.String) {
            _rowReports.Add((OfficeDiagnosticCodes.NullAsEmpty, OfficeConversionLossKind.Approximation,
                "Nulls in text columns became empty fields, which read back as empty strings.", node.Path));
        }

        _kinds[column] = kind;
        if (kind == DataScalarKind.Null) {
            _cellLengths[column] = -1;
            return;
        }

        if (kind == DataScalarKind.String && utf8.Length > 0 && MaybeTyped(utf8[0]) && LooksTyped(utf8)) {
            _rowReports.Add((OfficeDiagnosticCodes.ScalarTypeDropped, OfficeConversionLossKind.Approximation,
                "Text values that look like numbers, booleans or dates read back as those types from CSV.", Location(token)));
        }

        utf8.CopyTo(ReserveCell(column, utf8.Length));
        _cellLengths[column] = utf8.Length;
    }

    private static bool LooksTyped(ReadOnlySpan<byte> utf8) {
        char[] chars = ArrayPool<char>.Shared.Rent(utf8.Length);
        try {
            int length = Encoding.UTF8.GetChars(utf8, chars);
            return TableTypes.LooksTyped(new ReadOnlySpan<char>(chars, 0, length));
        } finally {
            ArrayPool<char>.Shared.Return(chars);
        }
    }

    // Names and values unescape into separate buffers, because a name is still in use while its value is read.
    private ReadOnlySpan<byte> Unescaped(ref Utf8JsonReader reader, bool names) {
        if (!reader.ValueIsEscaped) return reader.ValueSpan;
        int room = reader.ValueSpan.Length;
        byte[] scratch = names ? _nameScratch : _valueScratch;
        if (scratch.Length < room) {
            scratch = new byte[Math.Max(room, scratch.Length * 2)];
            if (names) _nameScratch = scratch;
            else _valueScratch = scratch;
        }

        return new ReadOnlySpan<byte>(scratch, 0, reader.CopyString(scratch));
    }

    private InvalidDataException Duplicate(ReadOnlySpan<byte> name, long token) =>
        new("The key '" + Encoding.UTF8.GetString(name) + "' appears twice in one object at " + Location(token) + ".");
}
#endif
