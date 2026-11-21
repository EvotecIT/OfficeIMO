using System.Text.Encodings.Web;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Writes JSON or JSON Lines with <see cref="Utf8JsonWriter"/>. Several documents become one array (JSON) or one line
/// each (JSON Lines); a single array document written as JSON Lines becomes one line per item.
/// </summary>
internal sealed class JsonEventWriter : DataEventWriter {
    private static readonly byte[] NewLine = { (byte)'\n' };
    private readonly Stream _stream;
    private readonly Utf8JsonWriter _writer;
    private readonly bool _jsonLines;
    private readonly bool _wrapDocuments;
    private readonly bool _reportCombined;
    private readonly bool _explodeCandidate;
    private bool _exploding;
    private int _depth;
    private int _documents;

    /// <param name="stream">The UTF-8 output.</param>
    /// <param name="expectedDocuments">How many documents the source holds; null when it is a list of records (JSON Lines).</param>
    /// <param name="context">The conversion context.</param>
    /// <param name="reportCombined">Whether joining several documents into one array is a reported change (YAML streams).</param>
    public JsonEventWriter(Stream stream, long? expectedDocuments, DataConversionContext context, bool reportCombined) : base(context) {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _jsonLines = context.Options.JsonLines;
        JsonWriterOptions options = context.Options.JsonWriter ?? new JsonWriterOptions {
            Indented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };
        if (_jsonLines) options.Indented = false;
        _writer = new Utf8JsonWriter(stream, options);
        _wrapDocuments = !_jsonLines && expectedDocuments is null or > 1;
        _reportCombined = reportCombined && expectedDocuments > 1;
        _explodeCandidate = _jsonLines && expectedDocuments == 1;
    }

    private int LineDepth => _exploding ? 1 : 0;

    public override void WriteStartDocument() {
        if (_wrapDocuments && _documents == 0) _writer.WriteStartArray();
        if (_reportCombined && _documents == 1) {
            Context.Report(OfficeDiagnosticCodes.YamlDocumentsCombined, OfficeConversionLossKind.Approximation,
                "A multi-document YAML stream became one JSON array of its documents; write JSON Lines to keep one document per line.");
        }

        _documents++;
        Context.ItemsWritten++;
    }

    public override void WriteEndDocument() {
    }

    public override void WriteStartObject() {
        _writer.WriteStartObject();
        _depth++;
    }

    public override void WriteEndObject() {
        _writer.WriteEndObject();
        _depth--;
        EndValue();
    }

    public override void WriteStartArray() {
        if (_explodeCandidate && _depth == 0) {
            _exploding = true;
            _depth++;
            return;
        }

        _writer.WriteStartArray();
        _depth++;
    }

    public override void WriteEndArray() {
        _depth--;
        if (_exploding && _depth == 0) return;
        _writer.WriteEndArray();
        EndValue();
    }

    public override void WritePropertyName(ReadOnlySpan<char> name) => _writer.WritePropertyName(name);

    public override void WriteScalar(DataScalarKind kind, ReadOnlySpan<char> text) {
        switch (kind) {
            case DataScalarKind.Null:
                _writer.WriteNullValue();
                break;
            case DataScalarKind.Boolean:
                _writer.WriteBooleanValue(text.Length == 4);
                break;
            case DataScalarKind.Integer:
                _writer.WriteRawValue(text, skipInputValidation: true);
                break;
            case DataScalarKind.Float:
                if (DataScalarKinds.IsNonFinite(text)) {
                    Context.Report(OfficeDiagnosticCodes.NonFiniteNumber, OfficeConversionLossKind.Approximation,
                        "JSON has no NaN or infinity; those numbers were written as the strings \"NaN\", \"Infinity\" and \"-Infinity\".");
                    _writer.WriteStringValue(text);
                } else {
                    _writer.WriteRawValue(text, skipInputValidation: true);
                }

                break;
            case DataScalarKind.String:
                _writer.WriteStringValue(text);
                break;
            default:
                Context.Report(OfficeDiagnosticCodes.DatetimeAsString, OfficeConversionLossKind.Approximation,
                    "JSON has no date or time type; date-times were written as RFC 3339 strings.");
                _writer.WriteStringValue(text);
                break;
        }

        EndValue();
    }

    private void EndValue() {
        if (_jsonLines && _depth == LineDepth) {
            _writer.Flush();
            _stream.Write(NewLine, 0, 1);
            _writer.Reset();
        }
    }

    public override void Complete() {
        if (_wrapDocuments) {
            if (_documents == 0) _writer.WriteStartArray();
            _writer.WriteEndArray();
        } else if (_documents == 0 && !_jsonLines) {
            _writer.WriteNullValue();
        }

        _writer.Flush();
        if (!_jsonLines) _stream.Write(NewLine, 0, 1);
        _stream.Flush();
    }

    public override void Dispose() => _writer.Dispose();
}
