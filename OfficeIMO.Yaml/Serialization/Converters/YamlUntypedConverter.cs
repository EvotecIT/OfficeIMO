using System.Collections;
using System.Numerics;

namespace OfficeIMO.Yaml;

/// <summary>
/// <see cref="object"/> values. Reading builds plain .NET values: mappings become <see cref="Dictionary{TKey, TValue}"/>
/// of <see cref="string"/> to <see cref="object"/> in file order, sequences <see cref="List{T}"/> of <see cref="object"/>,
/// and scalars resolve through the document's schema to <see langword="null"/>, <see cref="bool"/>, <see cref="long"/>
/// (or <see cref="BigInteger"/> beyond 64 bits), <see cref="double"/> or <see cref="string"/>; timestamps and custom
/// tags stay strings. Writing uses the runtime type: the common .NET values, dictionaries and enumerables directly, and
/// other types through the options' metadata.
/// </summary>
internal sealed class YamlUntypedConverter : YamlConverter<object> {
    public static YamlUntypedConverter Instance { get; } = new YamlUntypedConverter();

    // Booleans are read as one of two boxes instead of a new box per value.
    private static readonly object True = true;
    private static readonly object False = false;

    private YamlTypeInfo<Dictionary<string, object?>>? _mapping;
    private YamlTypeInfo<List<object?>>? _sequence;

    public override void Configure(YamlSerializerOptions options) {
        _mapping = (YamlTypeInfo<Dictionary<string, object?>>)options.GetTypeInfo(typeof(Dictionary<string, object?>));
        _sequence = (YamlTypeInfo<List<object?>>)options.GetTypeInfo(typeof(List<object?>));
    }

    public override object Read(YamlEventCursor cursor, YamlReadState state) {
        switch (cursor.TokenType) {
            case YamlTokenType.MappingStart:
                return _mapping!.ReadCore(cursor, state);
            case YamlTokenType.SequenceStart:
                return _sequence!.ReadCore(cursor, state);
            case YamlTokenType.Scalar:
                return ReadScalar(cursor)!;
            default:
                throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(object));
        }
    }

    private static object? ReadScalar(YamlEventCursor cursor) {
        ReadOnlySpan<char> text = cursor.ScalarSpan;
        YamlSchema schema = cursor.Schema;
        switch (cursor.ResolveKind()) {
            case YamlScalarKind.Null:
                return null;
            case YamlScalarKind.Boolean:
                if (YamlSchemaResolver.TryParseBoolean(text, schema, out bool flag)) return flag ? True : False;
                break;
            case YamlScalarKind.Integer:
                if (YamlSchemaResolver.TryParseInt64(text, schema, out long integer)) return integer;
                if (YamlBigIntegerConverter.TryParse(text, out BigInteger big)) return big;
                if (YamlSchemaResolver.TryParseDouble(text, schema, out double approximate)) return approximate;
                break;
            case YamlScalarKind.Float:
                if (YamlSchemaResolver.TryParseDouble(text, schema, out double number)) return number;
                break;
        }
        return cursor.GetString();
    }

    public override void Write(YamlWriter writer, object value, YamlWriteState state) {
        switch (value) {
            case string text: writer.WriteScalar(text); return;
            // The shapes reading produces come first, enumerated without boxing their enumerators; plain entries and
            // items go out in one step, as generated metadata writes them.
            case Dictionary<string, object?> mapping:
                if (!state.Begin(writer, value, typeof(object))) return;
                writer.WriteStartMapping();
                foreach (KeyValuePair<string, object?> entry in mapping) {
                    if (TryWritePlainEntry(writer, entry.Key, entry.Value)) continue;
                    writer.WriteKey(entry.Key);
                    WriteValue(writer, entry.Value, state);
                }
                writer.WriteEndMapping();
                state.End();
                return;
            case List<object?> list:
                if (!state.Begin(writer, value, typeof(object))) return;
                writer.WriteStartSequence();
                for (int i = 0; i < list.Count; i++) {
                    object? item = list[i];
                    if (TryWritePlainItem(writer, item)) continue;
                    WriteValue(writer, item, state);
                }
                writer.WriteEndSequence();
                state.End();
                return;
            case bool flag: writer.WriteScalar(flag); return;
            case int number: writer.WriteScalar(number); return;
            case long number: writer.WriteScalar(number); return;
            case double number: writer.WriteScalar(number); return;
            case short number: writer.WriteScalar(number); return;
            case byte number: writer.WriteScalar(number); return;
            case sbyte number: writer.WriteScalar(number); return;
            case ushort number: writer.WriteScalar(number); return;
            case uint number: writer.WriteScalar(number); return;
            case ulong number: YamlUInt64Converter.Instance.Write(writer, number, state); return;
            case float number: YamlSingleConverter.Instance.Write(writer, number, state); return;
            case decimal number: YamlDecimalConverter.Instance.Write(writer, number, state); return;
            case BigInteger number: YamlBigIntegerConverter.Instance.Write(writer, number, state); return;
            case char character: writer.WriteScalar(character.ToString()); return;
            case DateTime time: YamlDateTimeConverter.Instance.Write(writer, time, state); return;
            case DateTimeOffset time: YamlDateTimeOffsetConverter.Instance.Write(writer, time, state); return;
            case TimeSpan span: YamlTimeSpanConverter.Instance.Write(writer, span, state); return;
            case Guid guid: YamlGuidConverter.Instance.Write(writer, guid, state); return;
            case Uri uri: YamlUriConverter.Instance.Write(writer, uri, state); return;
            case byte[] bytes: YamlByteArrayConverter.Instance.Write(writer, bytes, state); return;
#if NET6_0_OR_GREATER
            case DateOnly date: YamlDateOnlyConverter.Instance.Write(writer, date, state); return;
            case TimeOnly time: YamlTimeOnlyConverter.Instance.Write(writer, time, state); return;
#endif
            case Enum member: writer.WriteScalar(member.ToString()); return;
            case IEnumerable<KeyValuePair<string, object?>> entries:
                if (!state.Begin(writer, value, typeof(object))) return;
                writer.WriteStartMapping();
                foreach (KeyValuePair<string, object?> entry in entries) {
                    writer.WriteKey(entry.Key);
                    WriteValue(writer, entry.Value, state);
                }
                writer.WriteEndMapping();
                state.End();
                return;
            case IDictionary dictionary:
                if (!state.Begin(writer, value, typeof(object))) return;
                writer.WriteStartMapping();
                foreach (DictionaryEntry entry in dictionary) {
                    if (entry.Key is string key) writer.WriteKey(key);
                    else WriteValue(writer, entry.Key, state);
                    WriteValue(writer, entry.Value, state);
                }
                writer.WriteEndMapping();
                state.End();
                return;
            case IEnumerable items:
                if (!state.Begin(writer, value, typeof(object))) return;
                writer.WriteStartSequence();
                foreach (object? item in items) WriteValue(writer, item, state);
                writer.WriteEndSequence();
                state.End();
                return;
        }
        state.Options.GetTypeInfo(value.GetType()).WriteBoxed(writer, value, state);
    }

    private static bool TryWritePlainEntry(YamlWriter writer, string key, object? value) => value switch {
        string text => writer.TryWriteStringEntry(key, text, checkKey: true),
        long number => YamlWriter.IsPlainKey(key) && writer.TryWritePlainEntry(key, number),
        bool flag => YamlWriter.IsPlainKey(key) && writer.TryWritePlainEntry(key, flag),
        double number => YamlWriter.IsPlainKey(key) && writer.TryWritePlainEntry(key, number),
        null => YamlWriter.IsPlainKey(key) && writer.TryWriteNullEntry(key),
        _ => false
    };

    private static bool TryWritePlainItem(YamlWriter writer, object? item) => item switch {
        string text => writer.TryWriteStringItem(text),
        long number => writer.TryWritePlainItem(number),
        bool flag => writer.TryWritePlainItem(flag),
        double number => writer.TryWritePlainItem(number),
        null => writer.TryWriteNullItem(),
        _ => false
    };

    private void WriteValue(YamlWriter writer, object? value, YamlWriteState state) {
        if (value is null) writer.WriteNull();
        else Write(writer, value, state);
    }

    public override void Collect(object value, YamlReferenceCollector collector) {
        switch (value) {
            case string:
                return;
            case IEnumerable<KeyValuePair<string, object?>> entries:
                if (!collector.Visit(value)) return;
                foreach (KeyValuePair<string, object?> entry in entries) {
                    if (entry.Value is not null) Collect(entry.Value, collector);
                }
                return;
            case IDictionary dictionary:
                if (!collector.Visit(value)) return;
                foreach (DictionaryEntry entry in dictionary) {
                    if (entry.Value is not null) Collect(entry.Value, collector);
                }
                return;
            case IEnumerable items:
                if (!collector.Visit(value)) return;
                foreach (object? item in items) {
                    if (item is not null) Collect(item, collector);
                }
                return;
        }
        if (value.GetType().IsValueType || value is Uri || value is byte[]) return;
        // Other objects are visited through their own metadata, as writing them does.
        collector.Options.GetTypeInfo(value.GetType()).CollectBoxed(value, collector);
    }
}
