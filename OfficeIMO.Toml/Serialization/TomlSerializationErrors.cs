namespace OfficeIMO.Toml;

/// <summary>
/// Builds the exceptions the serializer throws. Reading errors carry the position of the offending node and its path,
/// such as <c>tool.ruff["line-length"]</c> or <c>package[3].name</c>, in the grammar <see cref="TomlNode.Select"/> takes.
/// </summary>
internal static class TomlSerializationErrors {
    public static TomlException TypeMismatch(TomlNode node, Type type, string? detail = null) {
        string message = "Cannot read " + Describe(node) + " as " + TypeName(type) + At(node) + (detail is null ? "." : ": " + detail + ".");
        return new TomlException(TomlErrorCodes.TypeMismatch, message, PositionOf(node));
    }

    public static TomlException MissingRequired(TomlNode table, Type type, string key) =>
        new TomlException(TomlErrorCodes.MissingRequiredMember,
            "The table for " + TypeName(type) + At(table) + " has no '" + key + "' key, which is required.", PositionOf(table));

    public static TomlException UnknownDiscriminator(TomlNode node, Type type, string? discriminator) =>
        new TomlException(TomlErrorCodes.UnknownDiscriminator,
            discriminator is null
                ? "The table" + At(node) + " has no discriminator, and " + TypeName(type) + " cannot be created without one."
                : "The discriminator '" + Shorten(discriminator) + "'" + At(node) + " is not registered for " + TypeName(type) + ".",
            PositionOf(node));

    public static TomlException Cycle(Type type) =>
        new TomlException(TomlErrorCodes.ReferenceCycle, "A " + TypeName(type) + " contains itself, which TOML cannot write.", default);

    public static TomlException NullValue(string key) =>
        new TomlException(TomlErrorCodes.UnrepresentableValue,
            "The value of '" + key + "' is null, which TOML cannot hold. Set TomlSerializerOptions.NullHandling to Omit to leave the key out.", default);

    public static TomlException NullItem(Type type) =>
        new TomlException(TomlErrorCodes.UnrepresentableValue,
            "A " + TypeName(type) + " has a null item, which a TOML array cannot hold.", default);

    public static TomlException OutOfRange(string value) =>
        new TomlException(TomlErrorCodes.UnrepresentableValue, "The integer " + value + " is outside the 64-bit signed range of TOML integers.", default);

    public static NotSupportedException NotATable(Type type) =>
        new NotSupportedException("A TOML document is a table, so its root must be an object or a dictionary; " + TypeName(type) + " is not.");

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

    /// <summary>The node's position, or for a table without its own text (implicit or dotted) the first position inside it.</summary>
    public static TomlPosition PositionOf(TomlNode node) {
        for (TomlNode? current = node; current is not null; current = current.Parent) {
            TomlPosition start = current.Start;
            if (start.Line > 0) return start;
            TomlPosition inside = FirstInside(current);
            if (inside.Line > 0) return inside;
        }
        return default;
    }

    private static TomlPosition FirstInside(TomlNode node) {
        switch (node) {
            case TomlTable table:
                for (int i = 0; i < table._count; i++) {
                    TomlPosition position = table._nodes[i].Start;
                    if (position.Line > 0) return position;
                }
                break;
            case TomlTableArray tables:
                foreach (TomlTable table in tables) {
                    if (table.Start.Line > 0) return table.Start;
                }
                break;
        }
        return default;
    }

    private static string Describe(TomlNode node) => node switch {
        TomlTable => "a table",
        TomlTableArray => "an array of tables",
        TomlArray => "an array",
        TomlValue value when value.Kind == TomlNodeKind.String => "the string '" + Shorten(value.AsString()) + "'",
        TomlValue value => "the " + Kind(value.Kind) + " " + Shorten(value.AsString()),
        _ => "a value"
    };

    private static string Kind(TomlNodeKind kind) => kind switch {
        TomlNodeKind.Integer => "integer",
        TomlNodeKind.Float => "float",
        TomlNodeKind.Boolean => "boolean",
        _ => "date-time"
    };

    private static string At(TomlNode node) {
        string path = PathOf(node);
        return path.Length == 0 ? string.Empty : " at " + path;
    }

    /// <summary>The node's path from the root, in the grammar of <see cref="TomlNode.Select"/>.</summary>
    internal static string PathOf(TomlNode node) {
        var steps = new List<string>();
        for (TomlNode? current = node; current?.Parent is TomlNode parent; current = parent) {
            switch (parent) {
                case TomlTable:
                    string key = current.Key!;
                    steps.Add(IsPlainKey(key) ? "." + key : "[\"" + Escape(key) + "\"]");
                    break;
                case TomlArray array:
                    steps.Add("[" + IndexIn(array, current).ToString(CultureInfo.InvariantCulture) + "]");
                    break;
                case TomlTableArray tables:
                    steps.Add("[" + IndexIn(tables, current).ToString(CultureInfo.InvariantCulture) + "]");
                    // The array of tables sits under its key in its own parent table.
                    break;
            }
        }
        if (steps.Count == 0) return string.Empty;
        steps.Reverse();
        string path = string.Concat(steps);
        return path[0] == '.' ? path.Substring(1) : path;
    }

    private static int IndexIn(IEnumerable<TomlNode> items, TomlNode item) {
        int index = 0;
        foreach (TomlNode candidate in items) {
            if (ReferenceEquals(candidate, item)) return index;
            index++;
        }
        return -1;
    }

    private static bool IsPlainKey(string key) {
        if (key.Length == 0) return false;
        foreach (char ch in key) {
            if (ch == '.' || ch == '[' || ch == ']' || ch == '"' || ch == '\\' || char.IsWhiteSpace(ch) || char.IsControl(ch)) return false;
        }
        return true;
    }

    private static string Escape(string key) {
        var text = new StringBuilder(key.Length + 2);
        foreach (char ch in key) {
            switch (ch) {
                case '"': text.Append("\\\""); break;
                case '\\': text.Append("\\\\"); break;
                case '\n': text.Append("\\n"); break;
                case '\r': text.Append("\\r"); break;
                case '\t': text.Append("\\t"); break;
                default:
                    if (char.IsControl(ch)) text.Append("\\u").Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                    else text.Append(ch);
                    break;
            }
        }
        return text.ToString();
    }

    private static string Shorten(string text) => text.Length <= 40 ? text : text.Substring(0, 37) + "...";
}
