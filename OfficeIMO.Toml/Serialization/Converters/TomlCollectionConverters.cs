#if NET8_0_OR_GREATER
using System.Runtime.InteropServices;
#endif

namespace OfficeIMO.Toml;

/// <summary>
/// A collection read from an array or an array of tables into a <see cref="List{T}"/> and turned into the target
/// collection by a finishing function. A non-empty collection of tables is written as <c>[[key]]</c> sections; any other
/// collection, and every collection inside an array or inline table, as an array.
/// </summary>
internal sealed class TomlCollectionConverter<TCollection, TElement> : TomlConverter<TCollection> where TCollection : IEnumerable<TElement> {
    private readonly Func<List<TElement>, TCollection> _finish;
    private readonly bool _listIsResult;
    private TomlTypeInfo<TElement>? _element;
    private bool _elementTables;
    private bool _elementDynamic;

    public TomlCollectionConverter(Func<List<TElement>, TCollection> finish) {
        _finish = finish;
        _listIsResult = typeof(TCollection).IsAssignableFrom(typeof(List<TElement>));
    }

    public override TomlShape Shape => TomlShape.Array;

    public override bool HasDynamicShape => _elementTables || _elementDynamic;

    public override void Configure(TomlSerializerOptions options) {
        _element = (TomlTypeInfo<TElement>)options.GetTypeInfo(typeof(TElement));
        _elementDynamic = _element.Converter.HasDynamicShape;
        _elementTables = !_elementDynamic && _element.StaticShape == TomlShape.Table;
    }

    public override TomlShape GetShape(TCollection value) {
        TomlTypeInfo<TElement> element = _element!;
        bool any = false;
        if (_elementTables) {
            if (value is ICollection<TElement> collection) return collection.Count > 0 ? TomlShape.TableArray : TomlShape.Array;
            foreach (TElement _ in value) return TomlShape.TableArray;
            return TomlShape.Array;
        }
        foreach (TElement item in value) {
            if (item is null || element.GetShape(item) != TomlShape.Table) return TomlShape.Array;
            any = true;
        }
        return any ? TomlShape.TableArray : TomlShape.Array;
    }

    public override TCollection Read(TomlNode node, TomlReadState state) {
        TomlTypeInfo<TElement> element = _element!;
        List<TElement> list;
        switch (node) {
            case TomlArray array:
                list = new List<TElement>(array.Count);
                for (int i = 0; i < array.Count; i++) list.Add(element.Read(array[i], state));
                break;
            case TomlTableArray tables:
                list = new List<TElement>(tables.Count);
                for (int i = 0; i < tables.Count; i++) list.Add(element.Read(tables[i], state));
                break;
            default:
                throw TomlSerializationErrors.TypeMismatch(node, typeof(TCollection));
        }
        return _listIsResult ? (TCollection)(object)list : _finish(list);
    }

    public override void Write(TomlEmitter emitter, TCollection value, TomlWriteState state) {
        TomlTypeInfo<TElement> element = _element!;
        emitter.BeginArray();
        if (value is List<TElement> list) {
            for (int i = 0; i < list.Count; i++) WriteItem(emitter, element, list[i], state);
        } else if (value is TElement[] array) {
            for (int i = 0; i < array.Length; i++) WriteItem(emitter, element, array[i], state);
        } else {
            foreach (TElement item in value) WriteItem(emitter, element, item, state);
        }
        emitter.EndArray();
    }

    private static void WriteItem(TomlEmitter emitter, TomlTypeInfo<TElement> element, TElement item, TomlWriteState state) {
        if (item is null) throw TomlSerializationErrors.NullItem(typeof(TCollection));
        element.Write(emitter, item, state);
    }

    public override void WriteTableArray(TomlEmitter emitter, TCollection value, TomlWriteState state) {
        TomlTypeInfo<TElement> element = _element!;
        foreach (TElement item in value) {
            if (item is null) throw TomlSerializationErrors.NullItem(typeof(TCollection));
            emitter.BeginTableArrayItem();
            element.WriteBody(emitter, item, state);
        }
    }
}

/// <summary>
/// A dictionary read from a table into a <see cref="Dictionary{TKey, TValue}"/> and turned into the target type by a
/// finishing function. Keys are strings, or scalars read from and written as their text. A value that is a table is
/// written as a <c>[key]</c> section after the other entries, unless the dictionary is inline.
/// </summary>
internal sealed class TomlDictionaryConverter<TDictionary, TKey, TValue> : TomlConverter<TDictionary>
    where TDictionary : IEnumerable<KeyValuePair<TKey, TValue>>
    where TKey : notnull {
    private readonly Func<Dictionary<TKey, TValue>, TDictionary> _finish;
    private readonly bool _dictionaryIsResult;
    private readonly bool _stringKeys = typeof(TKey) == typeof(string);
    private TomlTypeInfo<TKey>? _key;
    private TomlTypeInfo<TValue>? _value;
    private bool _valueIsDictionary;

    public TomlDictionaryConverter(Func<Dictionary<TKey, TValue>, TDictionary> finish) {
        _finish = finish;
        _dictionaryIsResult = typeof(TDictionary).IsAssignableFrom(typeof(Dictionary<TKey, TValue>));
    }

    public override TomlShape Shape => TomlShape.Table;

    public override bool IsDictionary => true;

    public override void Configure(TomlSerializerOptions options) {
        _key = (TomlTypeInfo<TKey>)options.GetTypeInfo(typeof(TKey));
        _value = (TomlTypeInfo<TValue>)options.GetTypeInfo(typeof(TValue));
        _valueIsDictionary = _value.Converter.IsDictionary;
    }

    public override TDictionary Read(TomlNode node, TomlReadState state) {
        if (node is not TomlTable table) throw TomlSerializationErrors.TypeMismatch(node, typeof(TDictionary));
        TomlTypeInfo<TValue> valueInfo = _value!;
        int count = table._count;
        var dictionary = new Dictionary<TKey, TValue>(count);
        string[] keys = table._keys;
        TomlNode[] nodes = table._nodes;
        for (int i = 0; i < count; i++) {
            string text = keys[i];
            TKey key;
            if (_stringKeys) {
                key = (TKey)(object)text;
            } else if (!_key!.Converter.TryReadKey(text, out key)) {
                throw TomlSerializationErrors.TypeMismatch(nodes[i], typeof(TKey), "the key '" + text + "' is not a " + TomlSerializationErrors.TypeName(typeof(TKey)));
            }
#if NET8_0_OR_GREATER
            // Keys are unique in a table, but two texts can read as one key, such as "1" and "01"; the last wins.
            CollectionsMarshal.GetValueRefOrAddDefault(dictionary, key, out _) = valueInfo.Read(nodes[i], state);
#else
            dictionary[key] = valueInfo.Read(nodes[i], state);
#endif
        }
        return _dictionaryIsResult ? (TDictionary)(object)dictionary : _finish(dictionary);
    }

    private string KeyText(TKey key) {
        if (_stringKeys) return (string)(object)key;
        return _key!.Converter.WriteKey(key) ?? throw new NotSupportedException("A " + TomlSerializationErrors.TypeName(typeof(TKey)) + " cannot be a TOML key.");
    }

    private bool Skip(TValue value, TomlWriteState state) {
        if (value is not null) return false;
        if (state.FailOnNull) throw TomlSerializationErrors.NullValue(TomlSerializationErrors.TypeName(typeof(TDictionary)) + " entry");
        return true;
    }

    public override void Write(TomlEmitter emitter, TDictionary value, TomlWriteState state) {
        TomlTypeInfo<TValue> valueInfo = _value!;
        emitter.BeginInlineTable();
        foreach (KeyValuePair<TKey, TValue> entry in value) {
            if (Skip(entry.Value, state)) continue;
            emitter.WriteKey(KeyText(entry.Key));
            valueInfo.Write(emitter, entry.Value, state);
        }
        emitter.EndInlineTable();
    }

    private TomlShape ShapeOf(TValue value, bool inlineEntries) {
        TomlShape shape = _value!.GetShape(value);
        if (!inlineEntries) return shape;
        if (shape == TomlShape.Table && !_valueIsDictionary) return TomlShape.InlineTable;
        return shape == TomlShape.TableArray ? TomlShape.Array : shape;
    }

    public override void WriteBody(TomlEmitter emitter, TDictionary value, TomlWriteState state) {
        TomlTypeInfo<TValue> valueInfo = _value!;
        bool inlineEntries = state.InlineEntries;
        state.InlineEntries = false;
        bool sections = false;
        foreach (KeyValuePair<TKey, TValue> entry in value) {
            if (Skip(entry.Value, state)) continue;
            if (ShapeOf(entry.Value, inlineEntries) is TomlShape.Table or TomlShape.TableArray) {
                sections = true;
                continue;
            }
            emitter.WriteKey(KeyText(entry.Key));
            valueInfo.Write(emitter, entry.Value, state);
            emitter.EndValue();
        }
        if (!sections) return;
        foreach (KeyValuePair<TKey, TValue> entry in value) {
            if (entry.Value is null) continue;
            TomlShape shape = ShapeOf(entry.Value, inlineEntries);
            if (shape == TomlShape.Table) {
                emitter.PushTable(KeyText(entry.Key));
                valueInfo.WriteBody(emitter, entry.Value, state);
                emitter.PopTable();
            } else if (shape == TomlShape.TableArray) {
                emitter.PushKey(KeyText(entry.Key));
                valueInfo.WriteTableArray(emitter, entry.Value, state);
                emitter.PopKey();
            }
        }
    }
}
