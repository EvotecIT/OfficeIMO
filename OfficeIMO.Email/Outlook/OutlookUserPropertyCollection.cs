namespace OfficeIMO.Email;

/// <summary>
/// Outlook-compatible user-defined fields backed by PidLidPropertyDefinitionStream and
/// string-named PS_PUBLIC_STRINGS values.
/// </summary>
public sealed class OutlookUserPropertyCollection : IReadOnlyCollection<OutlookUserProperty> {
    private readonly EmailDocument _document;

    internal OutlookUserPropertyCollection(EmailDocument document) {
        _document = document ?? throw new ArgumentNullException(nameof(document));
    }

    /// <summary>State of the item's complete field-definition stream.</summary>
    public OutlookUserPropertyDefinitionState DefinitionState => ParseDefinitions().State;

    /// <summary>Parse error for a corrupt or unsupported definition stream.</summary>
    public string? DefinitionError => ParseDefinitions().Error;

    /// <summary>Every decoded field definition, including preserved built-in bindings.</summary>
    public IReadOnlyList<OutlookUserPropertyDefinition> Definitions => ParseDefinitions().Definitions;

    /// <inheritdoc />
    public int Count => Snapshot().Count;

    /// <summary>Finds a user property case-insensitively.</summary>
    public OutlookUserProperty? Find(string name) {
        string validated = ValidateName(name);
        return Snapshot().FirstOrDefault(property =>
            string.Equals(property.Name, validated, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Attempts to read and convert a user property's value.</summary>
    public bool TryGetValue<T>(string name, out T? value) {
        OutlookUserProperty? property = Find(name);
        if (property?.Property == null) {
            value = default;
            return false;
        }
        if (property.FieldType == OutlookUserPropertyType.Duration && typeof(T) == typeof(TimeSpan) &&
            TryGetInt32(property.Value, out int minutes)) {
            value = (T)(object)TimeSpan.FromMinutes(minutes);
            return true;
        }
        if (typeof(T) == typeof(string[]) && TryGetStrings(property.Value, out string[] strings)) {
            value = (T)(object)strings;
            return true;
        }
        return MapiValueConverter.TryConvert(property.Value, out value);
    }

    /// <summary>Reads a typed value or returns the managed default when it is absent or incompatible.</summary>
    public T? GetValueOrDefault<T>(string name) => TryGetValue(name, out T? value) ? value : default;

    /// <summary>Creates or replaces a user property, inferring its Outlook field type.</summary>
    public OutlookUserProperty Set(string name, object value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        OutlookUserPropertyType type = InferType(value);
        return Set(name, value, type);
    }

    /// <summary>
    /// Creates or replaces a user property and its Outlook field definition. Existing definitions cannot be changed
    /// to a different known type unless <paramref name="allowTypeChange"/> is true.
    /// </summary>
    public OutlookUserProperty Set(string name, object value, OutlookUserPropertyType fieldType,
        bool allowTypeChange = false) {
        string validatedName = ValidateName(name);
        if (value == null) throw new ArgumentNullException(nameof(value));

        OutlookUserPropertyDefinitionCodec.ParseResult parsed = ParseDefinitions();
        OutlookUserPropertyDefinition? existingDefinition = parsed.Definitions.LastOrDefault(definition =>
            definition.IsCustom && string.Equals(definition.Name, validatedName, StringComparison.OrdinalIgnoreCase));
        if (!allowTypeChange && existingDefinition != null &&
            existingDefinition.FieldType != OutlookUserPropertyType.Unknown &&
            existingDefinition.FieldType != fieldType) {
            throw new InvalidOperationException(string.Concat("User property '", validatedName,
                "' is already defined as ", existingDefinition.FieldType.ToString(), "."));
        }

        NormalizedValue normalized = NormalizeValue(value, fieldType);
        byte[] definitionStream = OutlookUserPropertyDefinitionCodec.AddOrReplace(parsed, validatedName,
            fieldType, GetCodePage());

        // Definition encoding happens before either property is mutated so failures are transactional.
        _document.Mapi.Set(MapiKnownProperties.PidLid.PropertyDefinitionStream, definitionStream);
        var key = new MapiPropertyKey<object>(string.Concat("OutlookUserProperty:", validatedName),
            MapiPropertySets.PublicStrings, validatedName, normalized.WireType);
        _document.Mapi.SetValue(key, normalized.Value, normalized.WireType);
        return Find(validatedName)!;
    }

    /// <summary>Creates or replaces a multiple-string keyword field.</summary>
    public OutlookUserProperty SetKeywords(string name, IEnumerable<string> values,
        bool allowTypeChange = false) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        string[] snapshot = values.Select(value => value ?? string.Empty).ToArray();
        return Set(name, snapshot, OutlookUserPropertyType.Keywords, allowTypeChange);
    }

    /// <summary>Creates or replaces a duration field using its natural managed representation.</summary>
    public OutlookUserProperty SetDuration(string name, TimeSpan value, bool allowTypeChange = false) =>
        Set(name, value, OutlookUserPropertyType.Duration, allowTypeChange);

    /// <summary>Removes the named value and its custom field definition.</summary>
    /// <returns>True when either a value or definition was removed.</returns>
    public bool Remove(string name) {
        string validatedName = ValidateName(name);
        OutlookUserPropertyDefinitionCodec.ParseResult parsed = ParseDefinitions();
        byte[]? updatedStream = OutlookUserPropertyDefinitionCodec.Remove(parsed, validatedName);
        bool hadDefinition = parsed.Definitions.Any(definition => definition.IsCustom &&
            string.Equals(definition.Name, validatedName, StringComparison.OrdinalIgnoreCase));

        var valueKey = new MapiPropertyKey<object>(string.Concat("OutlookUserProperty:", validatedName),
            MapiPropertySets.PublicStrings, validatedName, MapiPropertyType.Unspecified);
        int removedValues = RemovePublicStringValue(valueKey.Name!);
        if (hadDefinition) {
            if (updatedStream == null || updatedStream.Length == 0) {
                _document.Mapi.Remove(MapiKnownProperties.PidLid.PropertyDefinitionStream);
            } else {
                _document.Mapi.Set(MapiKnownProperties.PidLid.PropertyDefinitionStream, updatedStream);
            }
        }
        return hadDefinition || removedValues != 0;
    }

    /// <inheritdoc />
    public IEnumerator<OutlookUserProperty> GetEnumerator() => Snapshot().GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

    private IReadOnlyList<OutlookUserProperty> Snapshot() {
        OutlookUserPropertyDefinitionCodec.ParseResult parsed = ParseDefinitions();
        var names = new List<string>();
        var definitions = new Dictionary<string, OutlookUserPropertyDefinition>(StringComparer.OrdinalIgnoreCase);
        foreach (OutlookUserPropertyDefinition definition in parsed.Definitions) {
            if (!definition.IsCustom) continue;
            definitions[definition.Name] = definition;
            if (!names.Contains(definition.Name, StringComparer.OrdinalIgnoreCase)) names.Add(definition.Name);
        }

        var values = new Dictionary<string, MapiProperty>(StringComparer.OrdinalIgnoreCase);
        foreach (MapiProperty property in _document.MapiProperties) {
            string? name = property.Name?.Name;
            if (property.Name?.PropertySet != MapiPropertySets.PublicStrings || name == null ||
                IsReservedPublicString(name)) continue;
            values[name] = property;
            if (!names.Contains(name, StringComparer.OrdinalIgnoreCase)) names.Add(name);
        }

        return names.Select(name => new OutlookUserProperty(name,
            values.TryGetValue(name, out MapiProperty? property) ? property : null,
            definitions.TryGetValue(name, out OutlookUserPropertyDefinition? definition) ? definition : null))
            .ToArray();
    }

    private OutlookUserPropertyDefinitionCodec.ParseResult ParseDefinitions() {
        byte[]? bytes = _document.Mapi.GetValueOrDefault(MapiKnownProperties.PidLid.PropertyDefinitionStream);
        return OutlookUserPropertyDefinitionCodec.Parse(bytes, GetCodePage());
    }

    private int RemovePublicStringValue(MapiNamedProperty name) {
        int removed = 0;
        for (int index = _document.MapiProperties.Count - 1; index >= 0; index--) {
            if (_document.MapiProperties[index].Name == null ||
                !_document.MapiProperties[index].Name!.Equals(name)) continue;
            _document.MapiProperties.RemoveAt(index);
            removed++;
        }
        return removed;
    }

    private int GetCodePage() => _document.OutlookCodePage.GetValueOrDefault(1252);

    private static bool IsReservedPublicString(string name) =>
        string.Equals(name, MapiKnownProperties.PidName.Keywords.Name!.Name,
            StringComparison.OrdinalIgnoreCase);

    private static string ValidateName(string name) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (name.Length == 0 || string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("A user property name cannot be empty.", nameof(name));
        }
        if (name.IndexOf('\0') >= 0) throw new ArgumentException("A user property name cannot contain a null character.", nameof(name));
        if (name.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(name));
        return name;
    }

    private static OutlookUserPropertyType InferType(object value) {
        if (value is string) return OutlookUserPropertyType.Text;
        if (value is bool) return OutlookUserPropertyType.Boolean;
        if (value is int || value is short || value is byte) return OutlookUserPropertyType.Integer;
        if (value is float || value is double) return OutlookUserPropertyType.Number;
        if (value is decimal) return OutlookUserPropertyType.Currency;
        if (value is DateTime || value is DateTimeOffset) return OutlookUserPropertyType.DateTime;
        if (value is TimeSpan) return OutlookUserPropertyType.Duration;
        if (value is IEnumerable<string>) return OutlookUserPropertyType.Keywords;
        throw new ArgumentException(string.Concat("Cannot infer an Outlook user property type from ",
            value.GetType().FullName, "."), nameof(value));
    }

    private static NormalizedValue NormalizeValue(object value, OutlookUserPropertyType fieldType) {
        switch (fieldType) {
            case OutlookUserPropertyType.Text:
                if (value is string text) return new NormalizedValue(MapiPropertyType.Unicode, text);
                break;
            case OutlookUserPropertyType.Number:
            case OutlookUserPropertyType.Percent:
                if (IsNumeric(value)) return new NormalizedValue(MapiPropertyType.Floating64,
                    Convert.ToDouble(value, CultureInfo.InvariantCulture));
                break;
            case OutlookUserPropertyType.Currency:
                if (IsNumeric(value)) return new NormalizedValue(MapiPropertyType.Currency,
                    Convert.ToDecimal(value, CultureInfo.InvariantCulture));
                break;
            case OutlookUserPropertyType.Boolean:
                if (value is bool boolean) return new NormalizedValue(MapiPropertyType.Boolean, boolean);
                break;
            case OutlookUserPropertyType.DateTime:
                if (value is DateTimeOffset dateTimeOffset) return new NormalizedValue(MapiPropertyType.Time, dateTimeOffset);
                if (value is DateTime dateTime) return new NormalizedValue(MapiPropertyType.Time,
                    new DateTimeOffset(dateTime));
                break;
            case OutlookUserPropertyType.Duration:
                if (value is TimeSpan duration) {
                    double totalMinutes = duration.TotalMinutes;
                    if (totalMinutes < int.MinValue || totalMinutes > int.MaxValue || totalMinutes != Math.Truncate(totalMinutes)) {
                        throw new ArgumentOutOfRangeException(nameof(value), "Outlook durations require a whole Int32 number of minutes.");
                    }
                    return new NormalizedValue(MapiPropertyType.Integer32, (int)totalMinutes);
                }
                if (TryGetInt32(value, out int minutes)) return new NormalizedValue(MapiPropertyType.Integer32, minutes);
                break;
            case OutlookUserPropertyType.Keywords:
                if (TryGetStrings(value, out string[] values)) {
                    return new NormalizedValue(MapiPropertyType.MultipleUnicode, values.Cast<object>().ToArray());
                }
                break;
            case OutlookUserPropertyType.Integer:
                if (TryGetInt32(value, out int integer)) return new NormalizedValue(MapiPropertyType.Integer32, integer);
                break;
        }
        throw new ArgumentException(string.Concat("Value type ", value.GetType().FullName,
            " is incompatible with Outlook field type ", fieldType.ToString(), "."), nameof(value));
    }

    private static bool IsNumeric(object value) => value is byte || value is short || value is int ||
        value is long || value is float || value is double || value is decimal;

    private static bool TryGetInt32(object? value, out int result) {
        try {
            if (value != null && IsNumeric(value)) {
                result = Convert.ToInt32(value, CultureInfo.InvariantCulture);
                return true;
            }
        } catch (Exception ex) when (ex is OverflowException || ex is FormatException || ex is InvalidCastException) { }
        result = 0;
        return false;
    }

    private static bool TryGetStrings(object? value, out string[] result) {
        if (value is string[] exact) {
            result = exact;
            return true;
        }
        if (value is IEnumerable<string> strings) {
            result = strings.ToArray();
            return true;
        }
        if (value is object[] objects && objects.All(item => item is string)) {
            result = objects.Cast<string>().ToArray();
            return true;
        }
        result = Array.Empty<string>();
        return false;
    }

    private sealed class NormalizedValue {
        internal NormalizedValue(MapiPropertyType wireType, object value) {
            WireType = wireType;
            Value = value;
        }
        internal MapiPropertyType WireType { get; }
        internal object Value { get; }
    }
}
