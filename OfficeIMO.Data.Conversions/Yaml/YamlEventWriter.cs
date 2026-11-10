using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Writes data as YAML through <see cref="YamlWriter"/>, which quotes any string a YAML 1.2 or 1.1 reader could take for
/// another type. Each source document is a YAML document, so JSON Lines becomes a multi-document stream. Numbers keep
/// their text, rewritten only where YAML 1.1 would not read it as the same number (<c>1e3</c> becomes <c>1.0e+3</c>).
/// Comments are carried.
/// </summary>
internal sealed class YamlEventWriter : DataEventWriter {
    private readonly YamlWriter _yaml;
    private int _documents;

    public YamlEventWriter(Stream output, DataConversionContext context) : base(context) {
        _yaml = new YamlWriter(output, context.Options.YamlWriter);
    }

    public override void WriteStartDocument() {
        if (_documents++ > 0) _yaml.WriteDocumentStart();
        Context.ItemsWritten++;
    }

    public override void WriteEndDocument() {
    }

    public override void WriteStartObject() => _yaml.WriteStartMapping();
    public override void WriteEndObject() => _yaml.WriteEndMapping();
    public override void WriteStartArray() => _yaml.WriteStartSequence();
    public override void WriteEndArray() => _yaml.WriteEndSequence();
    public override void WritePropertyName(ReadOnlySpan<char> name) => _yaml.WriteKey(name.ToString());
    public override void WriteComment(ReadOnlySpan<char> text) => _yaml.WriteComment(text.ToString());

    public override void WriteScalar(DataScalarKind kind, ReadOnlySpan<char> text) {
        switch (kind) {
            case DataScalarKind.Null:
                _yaml.WriteScalar((string?)null);
                break;
            case DataScalarKind.Boolean:
                _yaml.WriteScalar(text.Length == 4);
                break;
            case DataScalarKind.Integer:
                _yaml.WriteScalar(text.ToString(), YamlStyle.Plain);
                break;
            case DataScalarKind.Float:
                if (DataScalarKinds.IsNonFinite(text)) {
                    _yaml.WriteScalar(text[0] == 'N' ? double.NaN : text[0] == '-' ? double.NegativeInfinity : double.PositiveInfinity);
                } else {
                    _yaml.WriteScalar(Float(text), YamlStyle.Plain);
                }

                break;
            case DataScalarKind.String:
                _yaml.WriteScalar(text.ToString());
                break;
            default:
                Context.Report(OfficeDiagnosticCodes.DatetimeAsString, OfficeConversionLossKind.Approximation,
                    "The YAML 1.2 core schema has no date or time type; date-times were written as RFC 3339 strings.");
                _yaml.WriteScalar(text.ToString());
                break;
        }
    }

    // YAML 1.1 floats need a '.' in the mantissa and a signed exponent; 1.2 core reads that form the same way.
    internal static string Float(ReadOnlySpan<char> text) {
        int exponent = text.IndexOfAny('e', 'E');
        ReadOnlySpan<char> mantissa = exponent < 0 ? text : text.Slice(0, exponent);
        bool dot = mantissa.IndexOf('.') >= 0;
        if (exponent < 0) return dot ? text.ToString() : text.ToString() + ".0";
        ReadOnlySpan<char> power = text.Slice(exponent + 1);
        bool signed = power.Length > 0 && (power[0] == '+' || power[0] == '-');
        if (dot && signed) return text.ToString();
        return mantissa.ToString() + (dot ? string.Empty : ".0") + "e" + (signed ? string.Empty : "+") + power.ToString();
    }

    public override void Complete() => _yaml.Flush();

    public override void Dispose() => _yaml.Dispose();
}
