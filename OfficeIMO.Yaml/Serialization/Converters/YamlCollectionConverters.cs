namespace OfficeIMO.Yaml;

/// <summary>
/// A sequence read into a <see cref="List{T}"/> and turned into the target collection by a finishing function: the list
/// itself for <see cref="List{T}"/> and the list interfaces, an array, a set, or a collection filled with <c>Add</c>.
/// </summary>
internal sealed class YamlCollectionConverter<TCollection, TElement> : YamlConverter<TCollection> where TCollection : IEnumerable<TElement> {
    private readonly Func<List<TElement>, TCollection> _finish;
    private readonly bool _listIsResult;
    private YamlTypeInfo<TElement>? _element;

    public YamlCollectionConverter(Func<List<TElement>, TCollection> finish) {
        _finish = finish;
        _listIsResult = typeof(TCollection).IsAssignableFrom(typeof(List<TElement>));
    }

    public override bool IsContainer => true;

    public override void Configure(YamlSerializerOptions options) => _element = (YamlTypeInfo<TElement>)options.GetTypeInfo(typeof(TElement));

    public override TCollection Read(YamlEventCursor cursor, YamlReadState state) {
        if (cursor.TokenType != YamlTokenType.SequenceStart) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(TCollection));
        YamlTypeInfo<TElement> element = _element!;
        var list = new List<TElement>();
        if (_listIsResult && state.Preserve) state.TakePendingAnchor(list);
        state.PushIndex(0);
        while (true) {
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            if (cursor.TokenType == YamlTokenType.SequenceEnd) break;
            state.SetIndex(list.Count);
            list.Add(element.Read(cursor, state));
        }
        state.Pop();
        return _listIsResult ? (TCollection)(object)list : _finish(list);
    }

    public override void Write(YamlWriter writer, TCollection value, YamlWriteState state) {
        YamlTypeInfo<TElement> element = _element!;
        writer.WriteStartSequence();
        if (value is TElement[] array) {
            for (int i = 0; i < array.Length; i++) element.Write(writer, array[i], state);
        } else if (value is List<TElement> list) {
            for (int i = 0; i < list.Count; i++) element.Write(writer, list[i], state);
        } else {
            foreach (TElement item in value) element.Write(writer, item, state);
        }
        writer.WriteEndSequence();
    }

    public override void Collect(TCollection value, YamlReferenceCollector collector) {
        YamlTypeInfo<TElement> element = _element!;
        foreach (TElement item in value) element.Collect(item, collector);
    }
}

/// <summary>
/// A mapping read into a <see cref="Dictionary{TKey, TValue}"/> and turned into the target dictionary by a finishing
/// function. Keys are scalars read through their own type's metadata. Merge keys (<c>&lt;&lt;</c>) add the entries of
/// the merged mappings that the mapping does not set itself.
/// </summary>
internal sealed class YamlDictionaryConverter<TDictionary, TKey, TValue> : YamlConverter<TDictionary>
    where TDictionary : IEnumerable<KeyValuePair<TKey, TValue>>
    where TKey : notnull {
    private readonly Func<Dictionary<TKey, TValue>, TDictionary> _finish;
    private readonly bool _dictionaryIsResult;
    private YamlTypeInfo<TKey>? _key;
    private YamlTypeInfo<TValue>? _value;

    public YamlDictionaryConverter(Func<Dictionary<TKey, TValue>, TDictionary> finish) {
        _finish = finish;
        _dictionaryIsResult = typeof(TDictionary).IsAssignableFrom(typeof(Dictionary<TKey, TValue>));
    }

    public override bool IsContainer => true;

    public override void Configure(YamlSerializerOptions options) {
        _key = (YamlTypeInfo<TKey>)options.GetTypeInfo(typeof(TKey));
        _value = (YamlTypeInfo<TValue>)options.GetTypeInfo(typeof(TValue));
    }

    public override TDictionary Read(YamlEventCursor cursor, YamlReadState state) {
        if (cursor.TokenType != YamlTokenType.MappingStart) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(TDictionary));
        var dictionary = new Dictionary<TKey, TValue>();
        if (_dictionaryIsResult && state.Preserve) state.TakePendingAnchor(dictionary);
        HashSet<TKey>? merged = null;
        ReadEntries(cursor, state, dictionary, ref merged, merging: false);
        return _dictionaryIsResult ? (TDictionary)(object)dictionary : _finish(dictionary);
    }

    /// <summary>Reads entries up to the mapping's end. Merged entries never replace a key the mapping or an earlier merge set.</summary>
    private void ReadEntries(YamlEventCursor cursor, YamlReadState state, Dictionary<TKey, TValue> dictionary, ref HashSet<TKey>? merged, bool merging) {
        YamlTypeInfo<TKey> keyInfo = _key!;
        YamlTypeInfo<TValue> valueInfo = _value!;
        while (true) {
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            if (cursor.TokenType == YamlTokenType.MappingEnd) return;
            if (cursor.IsMergeKey) {
                if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
                ReadMerge(cursor, state, dictionary, ref merged);
                continue;
            }
            if (cursor.IsNullScalar) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(TKey), "a mapping key cannot be null here");
            string keyText = cursor.TokenType == YamlTokenType.Scalar ? cursor.GetString() : string.Empty;
            TKey key = keyInfo.Read(cursor, state);
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            state.PushKey(keyText);
            if (merging) {
                if (dictionary.ContainsKey(key)) {
                    cursor.Skip();
                } else {
                    dictionary.Add(key, valueInfo.Read(cursor, state));
                    (merged ??= new HashSet<TKey>(dictionary.Comparer)).Add(key);
                }
            } else if (dictionary.ContainsKey(key)) {
                // An explicit key replaces a merged value once; a second explicit occurrence is a duplicate.
                if (merged is null || !merged.Remove(key)) throw YamlSerializationErrors.DuplicateKey(cursor, state, keyText);
                dictionary[key] = valueInfo.Read(cursor, state);
            } else {
                dictionary.Add(key, valueInfo.Read(cursor, state));
            }
            state.Pop();
        }
    }

    private void ReadMerge(YamlEventCursor cursor, YamlReadState state, Dictionary<TKey, TValue> dictionary, ref HashSet<TKey>? merged) {
        if (cursor.TokenType == YamlTokenType.Alias) cursor.ExpandAlias();
        if (cursor.TokenType == YamlTokenType.MappingStart) {
            ReadEntries(cursor, state, dictionary, ref merged, merging: true);
            return;
        }
        if (cursor.TokenType != YamlTokenType.SequenceStart) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(TDictionary), "a merge key needs a mapping or a sequence of mappings");
        while (true) {
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            if (cursor.TokenType == YamlTokenType.SequenceEnd) return;
            if (cursor.TokenType == YamlTokenType.Alias) cursor.ExpandAlias();
            if (cursor.TokenType != YamlTokenType.MappingStart) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(TDictionary), "a merge key needs a mapping or a sequence of mappings");
            ReadEntries(cursor, state, dictionary, ref merged, merging: true);
        }
    }

    public override void Write(YamlWriter writer, TDictionary value, YamlWriteState state) {
        YamlTypeInfo<TKey> keyInfo = _key!;
        YamlTypeInfo<TValue> valueInfo = _value!;
        writer.WriteStartMapping();
        if (value is Dictionary<TKey, TValue> dictionary) {
            foreach (KeyValuePair<TKey, TValue> entry in dictionary) WriteEntry(writer, keyInfo, valueInfo, entry, state);
        } else {
            foreach (KeyValuePair<TKey, TValue> entry in value) WriteEntry(writer, keyInfo, valueInfo, entry, state);
        }
        writer.WriteEndMapping();
    }

    private static void WriteEntry(YamlWriter writer, YamlTypeInfo<TKey> keyInfo, YamlTypeInfo<TValue> valueInfo, KeyValuePair<TKey, TValue> entry, YamlWriteState state) {
        if (entry.Key is string text) writer.WriteKey(text);
        else keyInfo.Write(writer, entry.Key, state);
        valueInfo.Write(writer, entry.Value, state);
    }

    public override void Collect(TDictionary value, YamlReferenceCollector collector) {
        YamlTypeInfo<TValue> valueInfo = _value!;
        foreach (KeyValuePair<TKey, TValue> entry in value) valueInfo.Collect(entry.Value, collector);
    }
}
