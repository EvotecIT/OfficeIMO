using System.Diagnostics.CodeAnalysis;

// The overloads follow System.Text.Json, whose options parameters are optional on several overloads of one name.
#pragma warning disable RS0026

namespace OfficeIMO.Toml;

/// <summary>
/// Converts between .NET objects and TOML documents. The API follows <c>System.Text.Json</c> and
/// <c>YamlSerializer</c>, and reads the same <c>Office*</c> mapping attributes: overloads with
/// <see cref="TomlSerializerOptions"/> create metadata by reflection when no
/// <see cref="TomlSerializerOptions.TypeInfoResolver"/> is set, and overloads with <see cref="TomlTypeInfo{T}"/> or a
/// <see cref="TomlSerializerContext"/> use generated metadata, for trimmed and NativeAOT applications.
/// </summary>
/// <remarks>
/// <para>A TOML document is a table, so the root value is an object or a dictionary. Objects and dictionaries inside it
/// are written as <c>[table]</c> sections after the key/value pairs, non-empty collections of them as
/// <c>[[table]]</c> sections, and everything inside an array or inline table inline. <see cref="TomlInlineTableAttribute"/>
/// writes a type or property inline. TOML has no null: null values are left out (see
/// <see cref="TomlSerializerOptions.NullHandling"/>), so they read back as the property's default.</para>
/// <para>Reading loads the document with <see cref="TomlDocument"/>, so every TOML rule is checked, and maps its tables.
/// Values that do not fit their type fail with a <see cref="TomlException"/> whose <see cref="TomlException.Code"/> is
/// <see cref="TomlErrorCodes.TypeMismatch"/> and whose position and message point at the value.</para>
/// </remarks>
public static class TomlSerializer {
    private const string ReflectionSwitch = "OfficeIMO.Toml.TomlSerializer.IsReflectionEnabledByDefault";

    /// <summary>
    /// Gets whether overloads without generated metadata may use reflection. It is on unless the application turns it
    /// off; the package turns it off for <c>PublishAot</c> and <c>PublishTrimmed</c> applications through the
    /// <c>TomlSerializerIsReflectionEnabledByDefault</c> MSBuild property.
    /// </summary>
#if NET9_0_OR_GREATER
    [FeatureSwitchDefinition(ReflectionSwitch)]
#endif
    internal static bool IsReflectionEnabledByDefault { get; } = !AppContext.TryGetSwitch(ReflectionSwitch, out bool enabled) || enabled;

    private static readonly UTF8Encoding Utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    // ------------------------------------------------------------------ reflection-capable overloads

    /// <summary>Writes <paramref name="value"/> as a TOML document and returns the text.</summary>
    /// <exception cref="NotSupportedException"><typeparamref name="T"/> is not written as a table.</exception>
    /// <exception cref="TomlException">The value holds something TOML cannot write, such as a null array item or a cycle.</exception>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static string Serialize<T>(T value, TomlSerializerOptions? options = null) => Serialize(value, GetTypeInfo<T>(options));

    /// <summary>Writes <paramref name="value"/> as a TOML document and returns its UTF-8 bytes, without a byte-order mark.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static byte[] SerializeToUtf8Bytes<T>(T value, TomlSerializerOptions? options = null) => SerializeToUtf8Bytes(value, GetTypeInfo<T>(options));

    /// <summary>Writes <paramref name="value"/> as a UTF-8 TOML document to <paramref name="utf8Toml"/>, which stays open.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize<T>(Stream utf8Toml, T value, TomlSerializerOptions? options = null) => Serialize(utf8Toml, value, GetTypeInfo<T>(options));

    /// <summary>Writes <paramref name="value"/> as a TOML document to <paramref name="writer"/>, which stays open.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize<T>(TextWriter writer, T value, TomlSerializerOptions? options = null) => Serialize(writer, value, GetTypeInfo<T>(options));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> and returns the text.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static string Serialize(object value, Type inputType, TomlSerializerOptions? options = null) => WriteText(value, GetTypeInfo(inputType, options, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> and returns its UTF-8 bytes.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static byte[] SerializeToUtf8Bytes(object value, Type inputType, TomlSerializerOptions? options = null) => Utf8.GetBytes(WriteText(value, GetTypeInfo(inputType, options, value)));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> to a UTF-8 stream, which stays open.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize(Stream utf8Toml, object value, Type inputType, TomlSerializerOptions? options = null) => WriteStream(utf8Toml, value, GetTypeInfo(inputType, options, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> to a text writer, which stays open.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize(TextWriter writer, object value, Type inputType, TomlSerializerOptions? options = null) => WriteText(writer, value, GetTypeInfo(inputType, options, value));

    /// <summary>Reads a TOML document as <typeparamref name="T"/>.</summary>
    /// <exception cref="TomlException">The input is malformed, exceeds a limit or does not fit <typeparamref name="T"/>.</exception>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static T Deserialize<T>(string toml, TomlSerializerOptions? options = null) => Deserialize(toml, GetTypeInfo<T>(options));

    /// <summary>Reads a UTF-8 TOML document from a stream as <typeparamref name="T"/>. The stream stays open.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static T Deserialize<T>(Stream utf8Toml, TomlSerializerOptions? options = null) => Deserialize(utf8Toml, GetTypeInfo<T>(options));

    /// <summary>Reads a TOML document from a text reader as <typeparamref name="T"/>. The reader stays open.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static T Deserialize<T>(TextReader reader, TomlSerializerOptions? options = null) => Deserialize(reader, GetTypeInfo<T>(options));

    /// <summary>
    /// Reads a node of a loaded <see cref="TomlDocument"/> as <typeparamref name="T"/>, such as the
    /// <c>tool.ruff</c> table of a <c>pyproject.toml</c>.
    /// </summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static T Deserialize<T>(TomlNode node, TomlSerializerOptions? options = null) => Deserialize(node, GetTypeInfo<T>(options));

    /// <summary>Reads a TOML document as <paramref name="returnType"/>.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static object Deserialize(string toml, Type returnType, TomlSerializerOptions? options = null) {
        TomlTypeInfo info = Info(returnType, options);
        return ReadBoxed(Parse(toml, info), info);
    }

    /// <summary>Reads a UTF-8 TOML document from a stream as <paramref name="returnType"/>. The stream stays open.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static object Deserialize(Stream utf8Toml, Type returnType, TomlSerializerOptions? options = null) {
        TomlTypeInfo info = Info(returnType, options);
        return ReadBoxed(Load(utf8Toml, info), info);
    }

    /// <summary>Reads a TOML document from a text reader as <paramref name="returnType"/>. The reader stays open.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static object Deserialize(TextReader reader, Type returnType, TomlSerializerOptions? options = null) {
        TomlTypeInfo info = Info(returnType, options);
        return ReadBoxed(Parse(ReadAll(reader), info), info);
    }

    /// <summary>Reads a node of a loaded <see cref="TomlDocument"/> as <paramref name="returnType"/>.</summary>
    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    public static object Deserialize(TomlNode node, Type returnType, TomlSerializerOptions? options = null) => ReadBoxed(RequireNode(node), Info(returnType, options));

    // ------------------------------------------------------------------ generated metadata: TomlTypeInfo<T>

    /// <summary>Writes <paramref name="value"/> with generated metadata and returns the text.</summary>
    public static string Serialize<T>(T value, TomlTypeInfo<T> typeInfo) {
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        TomlEmitter emitter = Emit(value, typeInfo);
        string text = emitter.Text.ToString();
        emitter.Return();
        return text;
    }

    /// <summary>Writes <paramref name="value"/> with generated metadata and returns its UTF-8 bytes, without a byte-order mark.</summary>
    public static byte[] SerializeToUtf8Bytes<T>(T value, TomlTypeInfo<T> typeInfo) {
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        TomlEmitter emitter = Emit(value, typeInfo);
        byte[] bytes = ToUtf8(emitter.Text);
        emitter.Return();
        return bytes;
    }

    /// <summary>Writes <paramref name="value"/> with generated metadata to a UTF-8 stream, which stays open.</summary>
    public static void Serialize<T>(Stream utf8Toml, T value, TomlTypeInfo<T> typeInfo) {
        if (utf8Toml is null) throw new ArgumentNullException(nameof(utf8Toml));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        TomlEmitter emitter = Emit(value, typeInfo);
        WriteUtf8(utf8Toml, emitter.Text);
        emitter.Return();
    }

    /// <summary>Writes <paramref name="value"/> with generated metadata to a text writer, which stays open.</summary>
    public static void Serialize<T>(TextWriter writer, T value, TomlTypeInfo<T> typeInfo) {
        if (writer is null) throw new ArgumentNullException(nameof(writer));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        TomlEmitter emitter = Emit(value, typeInfo);
        Copy(writer, emitter.Text);
        emitter.Return();
    }

    /// <summary>Reads a TOML document with generated metadata.</summary>
    /// <exception cref="TomlException">The input is malformed, exceeds a limit or does not fit the type.</exception>
    public static T Deserialize<T>(string toml, TomlTypeInfo<T> typeInfo) {
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        return Read(Parse(toml, typeInfo), typeInfo);
    }

    /// <summary>Reads a UTF-8 TOML document from a stream with generated metadata. The stream stays open.</summary>
    public static T Deserialize<T>(Stream utf8Toml, TomlTypeInfo<T> typeInfo) {
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        return Read(Load(utf8Toml, typeInfo), typeInfo);
    }

    /// <summary>Reads a TOML document from a text reader with generated metadata. The reader stays open.</summary>
    public static T Deserialize<T>(TextReader reader, TomlTypeInfo<T> typeInfo) {
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        return Read(Parse(ReadAll(reader), typeInfo), typeInfo);
    }

    /// <summary>Reads a node of a loaded <see cref="TomlDocument"/> with generated metadata.</summary>
    public static T Deserialize<T>(TomlNode node, TomlTypeInfo<T> typeInfo) {
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        return Read(RequireNode(node), typeInfo);
    }

    // ------------------------------------------------------------------ generated metadata: context and Type

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> with the context's metadata and returns the text.</summary>
    public static string Serialize(object value, Type inputType, TomlSerializerContext context) => WriteText(value, GetTypeInfo(inputType, context, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> with the context's metadata and returns its UTF-8 bytes.</summary>
    public static byte[] SerializeToUtf8Bytes(object value, Type inputType, TomlSerializerContext context) => Utf8.GetBytes(WriteText(value, GetTypeInfo(inputType, context, value)));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> with the context's metadata to a UTF-8 stream, which stays open.</summary>
    public static void Serialize(Stream utf8Toml, object value, Type inputType, TomlSerializerContext context) => WriteStream(utf8Toml, value, GetTypeInfo(inputType, context, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> with the context's metadata to a text writer, which stays open.</summary>
    public static void Serialize(TextWriter writer, object value, Type inputType, TomlSerializerContext context) => WriteText(writer, value, GetTypeInfo(inputType, context, value));

    /// <summary>Reads a TOML document as <paramref name="returnType"/> with the context's metadata.</summary>
    public static object Deserialize(string toml, Type returnType, TomlSerializerContext context) {
        TomlTypeInfo info = GetTypeInfo(returnType, context, null);
        return ReadBoxed(Parse(toml, info), info);
    }

    /// <summary>Reads a UTF-8 TOML document from a stream as <paramref name="returnType"/> with the context's metadata. The stream stays open.</summary>
    public static object Deserialize(Stream utf8Toml, Type returnType, TomlSerializerContext context) {
        TomlTypeInfo info = GetTypeInfo(returnType, context, null);
        return ReadBoxed(Load(utf8Toml, info), info);
    }

    /// <summary>Reads a TOML document from a text reader as <paramref name="returnType"/> with the context's metadata. The reader stays open.</summary>
    public static object Deserialize(TextReader reader, Type returnType, TomlSerializerContext context) {
        TomlTypeInfo info = GetTypeInfo(returnType, context, null);
        return ReadBoxed(Parse(ReadAll(reader), info), info);
    }

    /// <summary>Reads a node of a loaded <see cref="TomlDocument"/> as <paramref name="returnType"/> with the context's metadata.</summary>
    public static object Deserialize(TomlNode node, Type returnType, TomlSerializerContext context) => ReadBoxed(RequireNode(node), GetTypeInfo(returnType, context, null));

    // ------------------------------------------------------------------ cores

    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    private static TomlTypeInfo<T> GetTypeInfo<T>(TomlSerializerOptions? options) => (TomlTypeInfo<T>)GetTypeInfo(typeof(T), options, null);

    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    private static TomlTypeInfo Info(Type type, TomlSerializerOptions? options) => GetTypeInfo(type, options, null);

    [RequiresUnreferencedCode(TomlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(TomlReflectionResolver.DynamicCodeMessage)]
    private static TomlTypeInfo GetTypeInfo(Type type, TomlSerializerOptions? options, object? value) {
        if (type is null) throw new ArgumentNullException(nameof(type));
        CheckValue(type, value);
        options ??= TomlSerializerOptions.Default;
        if (options.TypeInfoResolver is null && IsReflectionEnabledByDefault) options.UseReflectionFallback(TomlReflectionResolver.GetInstance());
        return options.GetTypeInfo(type);
    }

    private static TomlTypeInfo GetTypeInfo(Type type, TomlSerializerContext context, object? value) {
        if (type is null) throw new ArgumentNullException(nameof(type));
        if (context is null) throw new ArgumentNullException(nameof(context));
        CheckValue(type, value);
        return context.Options.GetTypeInfo(type);
    }

    private static void CheckValue(Type type, object? value) {
        if (value is not null && !type.IsInstanceOfType(value)) {
            throw new ArgumentException("The value is a " + TomlSerializationErrors.TypeName(value.GetType()) + ", not a " + TomlSerializationErrors.TypeName(type) + ".", nameof(value));
        }
    }

    /// <summary>Writes the document into a rented emitter, which the caller copies out and returns.</summary>
    private static TomlEmitter Emit<T>(T value, TomlTypeInfo<T> typeInfo) {
        if (value is null) throw new ArgumentNullException(nameof(value), "A TOML document cannot be null.");
        TomlEmitter emitter = TomlEmitter.Rent(typeInfo.Options.NewLine);
        try {
            if (typeInfo.GetShape(value) is not (TomlShape.Table or TomlShape.InlineTable)) throw TomlSerializationErrors.NotATable(typeof(T));
            typeInfo.WriteBody(emitter, value, new TomlWriteState(typeInfo.Options));
        } catch {
            emitter.Return();
            throw;
        }
        return emitter;
    }

    private static TomlEmitter EmitBoxed(object value, TomlTypeInfo typeInfo) {
        if (value is null) throw new ArgumentNullException(nameof(value), "A TOML document cannot be null.");
        TomlEmitter emitter = TomlEmitter.Rent(typeInfo.Options.NewLine);
        try {
            if (typeInfo.GetShapeBoxed(value) is not (TomlShape.Table or TomlShape.InlineTable)) throw TomlSerializationErrors.NotATable(typeInfo.Type);
            typeInfo.WriteBodyBoxed(emitter, value, new TomlWriteState(typeInfo.Options));
        } catch {
            emitter.Return();
            throw;
        }
        return emitter;
    }

    private static string WriteText(object value, TomlTypeInfo typeInfo) {
        TomlEmitter emitter = EmitBoxed(value, typeInfo);
        string text = emitter.Text.ToString();
        emitter.Return();
        return text;
    }

    private static void WriteStream(Stream utf8Toml, object value, TomlTypeInfo typeInfo) {
        if (utf8Toml is null) throw new ArgumentNullException(nameof(utf8Toml));
        TomlEmitter emitter = EmitBoxed(value, typeInfo);
        WriteUtf8(utf8Toml, emitter.Text);
        emitter.Return();
    }

    private static void WriteText(TextWriter writer, object value, TomlTypeInfo typeInfo) {
        if (writer is null) throw new ArgumentNullException(nameof(writer));
        TomlEmitter emitter = EmitBoxed(value, typeInfo);
        Copy(writer, emitter.Text);
        emitter.Return();
    }

    private static void Copy(TextWriter writer, StringBuilder text) {
#if NETCOREAPP3_0_OR_GREATER
        writer.Write(text);
#else
        writer.Write(text.ToString());
#endif
    }

    private static byte[] ToUtf8(StringBuilder text) {
        char[] chars = ArrayPool<char>.Shared.Rent(text.Length);
        try {
            text.CopyTo(0, chars, 0, text.Length);
            return Utf8.GetBytes(chars, 0, text.Length);
        } finally {
            ArrayPool<char>.Shared.Return(chars);
        }
    }

    private static void WriteUtf8(Stream stream, StringBuilder text) {
        char[] chars = ArrayPool<char>.Shared.Rent(text.Length);
        byte[] bytes = ArrayPool<byte>.Shared.Rent(Utf8.GetMaxByteCount(text.Length));
        try {
            text.CopyTo(0, chars, 0, text.Length);
            int count = Utf8.GetBytes(chars, 0, text.Length, bytes, 0);
            stream.Write(bytes, 0, count);
        } finally {
            ArrayPool<char>.Shared.Return(chars);
            ArrayPool<byte>.Shared.Return(bytes);
        }
    }

    private static TomlTable Parse(string toml, TomlTypeInfo typeInfo) {
        if (toml is null) throw new ArgumentNullException(nameof(toml));
        return TomlDocument.Parse(toml, typeInfo.Options.EffectiveLoadOptions).Root;
    }

    private static TomlTable Load(Stream utf8Toml, TomlTypeInfo typeInfo) {
        if (utf8Toml is null) throw new ArgumentNullException(nameof(utf8Toml));
        return TomlDocument.Load(utf8Toml, typeInfo.Options.EffectiveLoadOptions).Root;
    }

    private static string ReadAll(TextReader reader) {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        return reader.ReadToEnd();
    }

    private static TomlNode RequireNode(TomlNode node) => node ?? throw new ArgumentNullException(nameof(node));

    private static T Read<T>(TomlNode node, TomlTypeInfo<T> typeInfo) => typeInfo.Read(node, new TomlReadState(typeInfo.Options));

    private static object ReadBoxed(TomlNode node, TomlTypeInfo typeInfo) => typeInfo.ReadBoxed(node, new TomlReadState(typeInfo.Options))!;
}
