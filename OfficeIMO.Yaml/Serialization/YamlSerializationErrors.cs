namespace OfficeIMO.Yaml;

/// <summary>Builds the positioned exceptions the serializer throws for input that does not fit the target types.</summary>
internal static class YamlSerializationErrors {
    public static YamlException TypeMismatch(YamlEventCursor cursor, YamlReadState state, Type type, string? detail = null) {
        string found = cursor.TokenType switch {
            YamlTokenType.MappingStart => "a mapping",
            YamlTokenType.SequenceStart => "a sequence",
            YamlTokenType.Alias => "an alias",
            YamlTokenType.Scalar => cursor.IsNullScalar ? "null" : "the scalar '" + Describe(cursor.GetString()) + "'",
            _ => "the end of the input"
        };
        string message = "Cannot read " + found + " as " + TypeName(type) + At(state) + (detail is null ? "." : ": " + detail);
        return new YamlException(YamlErrorCodes.TypeMismatch, message, cursor.Position);
    }

    public static YamlException UnexpectedEnd(YamlEventCursor cursor) =>
        new YamlException(YamlErrorCodes.UnexpectedToken, "The YAML input ended inside a node.", cursor.Position);

    public static YamlException MissingRequired(YamlPosition position, YamlReadState state, Type type, string key) =>
        new YamlException(YamlErrorCodes.MissingRequiredMember,
            "The mapping for " + TypeName(type) + At(state) + " has no '" + key + "' key, which is required.", position);

    public static YamlException DuplicateKey(YamlEventCursor cursor, YamlReadState state, string key) =>
        new YamlException(YamlErrorCodes.DuplicateKey, "The mapping" + At(state) + " has the key '" + Describe(key) + "' more than once.", cursor.Position);

    public static YamlException UnknownDiscriminator(YamlEventCursor cursor, YamlReadState state, Type type, string? discriminator) =>
        new YamlException(YamlErrorCodes.UnknownDiscriminator,
            (discriminator is null
                ? "The value" + At(state) + " has no discriminator, and " + TypeName(type) + " cannot be created without one."
                : "The discriminator '" + Describe(discriminator) + "'" + At(state) + " is not registered for " + TypeName(type) + "."),
            cursor.Position);

    public static YamlException Cycle(Type type) =>
        new YamlException(YamlErrorCodes.ReferenceCycle,
            "A " + TypeName(type) + " contains itself. Set YamlSerializerOptions.ReferenceHandling to Preserve to write it with an anchor and alias.",
            default);

    /// <summary>Formats a type name with its generic arguments, as C# spells it.</summary>
    public static string TypeName(Type type) {
        Type? underlying = Nullable.GetUnderlyingType(type);
        if (underlying is not null) return TypeName(underlying) + "?";
        if (type.IsArray) return TypeName(type.GetElementType()!) + "[]";
        if (!type.IsGenericType) return type.Name;
        string name = type.Name;
        int tick = name.IndexOf('`');
        if (tick > 0) name = name.Substring(0, tick);
        Type[] arguments = type.GetGenericArguments();
        var text = new StringBuilder(name).Append('<');
        for (int i = 0; i < arguments.Length; i++) {
            if (i > 0) text.Append(", ");
            text.Append(TypeName(arguments[i]));
        }
        return text.Append('>').ToString();
    }

    private static string At(YamlReadState state) {
        string path = state.FormatPath();
        return path.Length == 0 ? string.Empty : " at " + path;
    }

    private static string Describe(string text) => text.Length <= 40 ? text : text.Substring(0, 37) + "...";
}
