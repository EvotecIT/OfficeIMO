namespace OfficeIMO.Toml;

/// <summary>
/// Values read as <see cref="object"/>: a table becomes a <c>Dictionary&lt;string, object?&gt;</c>, an array or array of
/// tables a <c>List&lt;object?&gt;</c>, and a value its decoded form: <see cref="string"/>, <see cref="long"/>,
/// <see cref="double"/>, <see cref="bool"/> or <see cref="TomlDateTime"/>, which keeps the date-time kind. Writing goes
/// by the runtime type: dictionaries are tables, other sequences arrays, and anything else uses that type's metadata.
/// </summary>
internal sealed class TomlUntypedConverter : TomlConverter<object> {
    private TomlSerializerOptions? _options;

    public override TomlShape Shape => TomlShape.Scalar;

    public override bool HasDynamicShape => true;

    public override bool TracksReferences => false;

    public override void Configure(TomlSerializerOptions options) => _options = options;

    public override object Read(TomlReadNode node, TomlReadState state) {
        switch (node.Kind) {
            case TomlNodeKind.Table: {
                var dictionary = new Dictionary<string, object?>(node.Count, StringComparer.Ordinal);
                foreach (TomlReadNode entry in node.GetChildren()) dictionary.Add(entry.Key!, Read(entry, state));
                return dictionary;
            }
            case TomlNodeKind.Array:
            case TomlNodeKind.TableArray: {
                var list = new List<object?>(node.Count);
                foreach (TomlReadNode item in node.GetChildren()) list.Add(Read(item, state));
                return list;
            }
            case TomlNodeKind.String:
                return TomlValueText.GetString(node);
            case TomlNodeKind.Integer:
                return TomlValueText.GetInt64(node);
            case TomlNodeKind.Float:
                return TomlValueText.GetDouble(node);
            case TomlNodeKind.Boolean:
                return TomlValueText.GetBoolean(node);
            default:
                return TomlValueText.GetDateTime(node);
        }
    }

    /// <summary>Sequences other than strings and byte arrays, which are written as one value.</summary>
    private static bool IsSequence(object value) => value is IEnumerable and not string and not byte[];

    private TomlTypeInfo RuntimeInfo(object value) => _options!.GetTypeInfo(value.GetType());

    public override TomlShape GetShape(object value) {
        if (value is IDictionary) return TomlShape.Table;
        if (IsSequence(value)) {
            bool any = false;
            foreach (object? item in (IEnumerable)value) {
                if (item is null || !IsTable(item)) return TomlShape.Array;
                any = true;
            }
            return any ? TomlShape.TableArray : TomlShape.Array;
        }
        TomlTypeInfo info = RuntimeInfo(value);
        return info.Type == typeof(object) ? TomlShape.Scalar : info.GetShapeBoxed(value);
    }

    /// <summary>Whether an item of a sequence is a table, without walking into nested sequences.</summary>
    private bool IsTable(object item) {
        if (item is IDictionary) return true;
        if (IsSequence(item)) return false;
        TomlTypeInfo info = RuntimeInfo(item);
        return info.Type != typeof(object) && info.GetShapeBoxed(item) == TomlShape.Table;
    }

    public override void Write(TomlEmitter emitter, object value, TomlWriteState state) {
        if (value is IDictionary dictionary) {
            state.Enter(value, value.GetType());
            emitter.BeginInlineTable();
            foreach (DictionaryEntry entry in dictionary) {
                if (Skip(entry.Value, state)) continue;
                emitter.WriteKey(KeyText(entry.Key));
                Write(emitter, entry.Value!, state);
            }
            emitter.EndInlineTable();
            state.Exit();
            return;
        }
        if (IsSequence(value)) {
            state.Enter(value, value.GetType());
            emitter.BeginArray();
            foreach (object? item in (IEnumerable)value) {
                if (item is null) throw TomlSerializationErrors.NullItem(value.GetType());
                Write(emitter, item, state);
            }
            emitter.EndArray();
            state.Exit();
            return;
        }
        TomlTypeInfo info = RuntimeInfo(value);
        if (info.Type == typeof(object)) throw new NotSupportedException("A plain object has no TOML representation.");
        info.WriteBoxed(emitter, value, state);
    }

    public override void WriteBody(TomlEmitter emitter, object value, TomlWriteState state) {
        if (value is not IDictionary dictionary) {
            RuntimeInfo(value).WriteBodyBoxed(emitter, value, state);
            return;
        }
        state.Enter(value, value.GetType());
        bool sections = false;
        foreach (DictionaryEntry entry in dictionary) {
            if (Skip(entry.Value, state)) continue;
            if (GetShape(entry.Value!) is TomlShape.Table or TomlShape.TableArray) {
                sections = true;
                continue;
            }
            emitter.WriteKey(KeyText(entry.Key));
            Write(emitter, entry.Value!, state);
            emitter.EndValue();
        }
        if (sections) {
            foreach (DictionaryEntry entry in dictionary) {
                if (entry.Value is null) continue;
                TomlShape shape = GetShape(entry.Value);
                if (shape == TomlShape.Table) {
                    emitter.PushTable(KeyText(entry.Key));
                    WriteBody(emitter, entry.Value, state);
                    emitter.PopTable();
                } else if (shape == TomlShape.TableArray) {
                    emitter.PushKey(KeyText(entry.Key));
                    WriteTableArray(emitter, entry.Value, state);
                    emitter.PopKey();
                }
            }
        }
        state.Exit();
    }

    public override void WriteTableArray(TomlEmitter emitter, object value, TomlWriteState state) {
        if (!IsSequence(value)) {
            RuntimeInfo(value).WriteTableArrayBoxed(emitter, value, state);
            return;
        }
        state.Enter(value, value.GetType());
        foreach (object? item in (IEnumerable)value) {
            emitter.BeginTableArrayItem();
            WriteBody(emitter, item!, state);
        }
        state.Exit();
    }

    private static bool Skip(object? value, TomlWriteState state) {
        if (value is not null) return false;
        if (state.FailOnNull) throw TomlSerializationErrors.NullValue("dictionary entry");
        return true;
    }

    private static string KeyText(object key) => key as string ?? Convert.ToString(key, CultureInfo.InvariantCulture) ?? string.Empty;
}
