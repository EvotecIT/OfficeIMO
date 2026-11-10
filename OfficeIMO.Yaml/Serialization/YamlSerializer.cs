using System.Diagnostics.CodeAnalysis;

// The overloads follow System.Text.Json, whose options parameters are optional on several overloads of one name.
#pragma warning disable RS0026

namespace OfficeIMO.Yaml;

/// <summary>
/// Converts between .NET objects and YAML, through <see cref="YamlReader"/> and <see cref="YamlWriter"/> directly.
/// The API follows <c>System.Text.Json</c>: overloads with <see cref="YamlSerializerOptions"/> create metadata by
/// reflection when no <see cref="YamlSerializerOptions.TypeInfoResolver"/> is set, and overloads with
/// <see cref="YamlTypeInfo{T}"/> or a <see cref="YamlSerializerContext"/> use generated metadata, for trimmed and
/// NativeAOT applications.
/// </summary>
/// <remarks>
/// <para>Reading resolves aliases by replaying their anchored node, within
/// <see cref="YamlReaderOptions.MaxAliasExpansion"/>, and applies <c>&lt;&lt;</c> merge keys. Values that do not fit
/// their type fail with a <see cref="YamlException"/> whose <see cref="YamlException.Code"/> is
/// <see cref="YamlErrorCodes.TypeMismatch"/> and whose position points at the node. Tags never select a .NET type:
/// only discriminators registered with <see cref="YamlDerivedTypeAttribute"/> do.</para>
/// <para>Text, stream and text-reader input must hold at most one document; read a multi-document stream with a
/// <see cref="YamlReader"/>, one call per document.</para>
/// </remarks>
public static class YamlSerializer {
    private const string ReflectionSwitch = "OfficeIMO.Yaml.YamlSerializer.IsReflectionEnabledByDefault";

    /// <summary>
    /// Gets whether overloads without generated metadata may use reflection. It is on unless the application turns it
    /// off; the package turns it off for <c>PublishAot</c> and <c>PublishTrimmed</c> applications through the
    /// <c>YamlSerializerIsReflectionEnabledByDefault</c> MSBuild property.
    /// </summary>
#if NET9_0_OR_GREATER
    [FeatureSwitchDefinition(ReflectionSwitch)]
#endif
    internal static bool IsReflectionEnabledByDefault { get; } = !AppContext.TryGetSwitch(ReflectionSwitch, out bool enabled) || enabled;

    // ------------------------------------------------------------------ reflection-capable overloads

    /// <summary>Writes <paramref name="value"/> as a YAML document and returns the text.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static string Serialize<T>(T value, YamlSerializerOptions? options = null) => Serialize(value, GetTypeInfo<T>(options));

    /// <summary>Writes <paramref name="value"/> as a UTF-8 YAML document to <paramref name="utf8Yaml"/>, which stays open.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize<T>(Stream utf8Yaml, T value, YamlSerializerOptions? options = null) => Serialize(utf8Yaml, value, GetTypeInfo<T>(options));

    /// <summary>Writes <paramref name="value"/> as a YAML document to <paramref name="writer"/>, which stays open.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize<T>(TextWriter writer, T value, YamlSerializerOptions? options = null) => Serialize(writer, value, GetTypeInfo<T>(options));

    /// <summary>
    /// Writes <paramref name="value"/> as one node at the writer's position: a document root, a mapping value or a
    /// sequence item. The writer's own layout settings apply.
    /// </summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize<T>(YamlWriter writer, T value, YamlSerializerOptions? options = null) => Serialize(writer, value, GetTypeInfo<T>(options));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> and returns the text.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static string Serialize(object? value, Type inputType, YamlSerializerOptions? options = null) => SerializeBoxed(value, GetTypeInfo(inputType, options, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> to a UTF-8 stream, which stays open.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize(Stream utf8Yaml, object? value, Type inputType, YamlSerializerOptions? options = null) => SerializeBoxed(utf8Yaml, value, GetTypeInfo(inputType, options, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> to a text writer, which stays open.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize(TextWriter writer, object? value, Type inputType, YamlSerializerOptions? options = null) => SerializeBoxed(writer, value, GetTypeInfo(inputType, options, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/>, as one node at the writer's position.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static void Serialize(YamlWriter writer, object? value, Type inputType, YamlSerializerOptions? options = null) => SerializeBoxed(writer, value, GetTypeInfo(inputType, options, value));

    /// <summary>
    /// Converts <paramref name="value"/> to a detached <see cref="YamlNode"/> tree, to insert into a
    /// <see cref="YamlDocument"/> next to hand-edited content.
    /// </summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static YamlNode SerializeToNode<T>(T value, YamlSerializerOptions? options = null) => SerializeToNode(value, GetTypeInfo<T>(options));

    /// <summary>Converts <paramref name="value"/> as <paramref name="inputType"/> to a detached <see cref="YamlNode"/> tree.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static YamlNode SerializeToNode(object? value, Type inputType, YamlSerializerOptions? options = null) => SerializeBoxedToNode(value, GetTypeInfo(inputType, options, value));

    /// <summary>Reads a YAML document as <typeparamref name="T"/>. Empty input reads as null.</summary>
    /// <exception cref="YamlException">The input is malformed, exceeds a limit or does not fit <typeparamref name="T"/>.</exception>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static T? Deserialize<T>(string yaml, YamlSerializerOptions? options = null) => Deserialize(yaml, GetTypeInfo<T>(options));

    /// <summary>Reads a YAML document from a stream as <typeparamref name="T"/>; the encoding is detected. The stream stays open.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static T? Deserialize<T>(Stream utf8Yaml, YamlSerializerOptions? options = null) => Deserialize(utf8Yaml, GetTypeInfo<T>(options));

    /// <summary>Reads a YAML document from a text reader as <typeparamref name="T"/>. The reader stays open.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static T? Deserialize<T>(TextReader reader, YamlSerializerOptions? options = null) => Deserialize(reader, GetTypeInfo<T>(options));

    /// <summary>
    /// Reads one value from a <see cref="YamlReader"/>. Between documents it reads the next document's root and stops on
    /// that document's end event, returning null at the end of the stream; on a node event it reads that node and stops
    /// on its last event. The reader's settings and limits apply.
    /// </summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static T? Deserialize<T>(YamlReader reader, YamlSerializerOptions? options = null) => Deserialize(reader, GetTypeInfo<T>(options));

    /// <summary>Reads a model node, such as a section of a loaded <see cref="YamlDocument"/>, as <typeparamref name="T"/>.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static T? Deserialize<T>(YamlNode node, YamlSerializerOptions? options = null) => Deserialize(node, GetTypeInfo<T>(options));

    /// <summary>Reads a YAML document as <paramref name="returnType"/>.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static object? Deserialize(string yaml, Type returnType, YamlSerializerOptions? options = null) => DeserializeBoxed(yaml, GetTypeInfo(returnType, options, null));

    /// <summary>Reads a YAML document from a stream as <paramref name="returnType"/>. The stream stays open.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static object? Deserialize(Stream utf8Yaml, Type returnType, YamlSerializerOptions? options = null) => DeserializeBoxed(utf8Yaml, GetTypeInfo(returnType, options, null));

    /// <summary>Reads a YAML document from a text reader as <paramref name="returnType"/>. The reader stays open.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static object? Deserialize(TextReader reader, Type returnType, YamlSerializerOptions? options = null) => DeserializeBoxed(reader, GetTypeInfo(returnType, options, null));

    /// <summary>Reads one value from a <see cref="YamlReader"/> as <paramref name="returnType"/>, as <see cref="Deserialize{T}(YamlReader, YamlSerializerOptions?)"/> does.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static object? Deserialize(YamlReader reader, Type returnType, YamlSerializerOptions? options = null) => DeserializeBoxed(reader, GetTypeInfo(returnType, options, null));

    /// <summary>Reads a model node as <paramref name="returnType"/>.</summary>
    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    public static object? Deserialize(YamlNode node, Type returnType, YamlSerializerOptions? options = null) => DeserializeBoxed(node, GetTypeInfo(returnType, options, null));

    // ------------------------------------------------------------------ generated metadata: YamlTypeInfo<T>

    /// <summary>Writes <paramref name="value"/> with generated metadata and returns the text.</summary>
    public static string Serialize<T>(T value, YamlTypeInfo<T> typeInfo) {
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        using var output = new YamlCharBuffer();
        using (var writer = new YamlWriter(output, typeInfo.Options.EffectiveWriterOptions)) Write(writer, value, typeInfo);
        return output.ToString();
    }

    /// <summary>Writes <paramref name="value"/> with generated metadata to a UTF-8 stream, which stays open.</summary>
    public static void Serialize<T>(Stream utf8Yaml, T value, YamlTypeInfo<T> typeInfo) {
        if (utf8Yaml is null) throw new ArgumentNullException(nameof(utf8Yaml));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        using var writer = new YamlWriter(utf8Yaml, typeInfo.Options.EffectiveWriterOptions);
        Write(writer, value, typeInfo);
    }

    /// <summary>Writes <paramref name="value"/> with generated metadata to a text writer, which stays open.</summary>
    public static void Serialize<T>(TextWriter writer, T value, YamlTypeInfo<T> typeInfo) {
        if (writer is null) throw new ArgumentNullException(nameof(writer));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        using var yaml = new YamlWriter(writer, typeInfo.Options.EffectiveWriterOptions);
        Write(yaml, value, typeInfo);
    }

    /// <summary>Writes <paramref name="value"/> with generated metadata, as one node at the writer's position.</summary>
    public static void Serialize<T>(YamlWriter writer, T value, YamlTypeInfo<T> typeInfo) {
        if (writer is null) throw new ArgumentNullException(nameof(writer));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        Write(writer, value, typeInfo);
    }

    /// <summary>Converts <paramref name="value"/> with generated metadata to a detached <see cref="YamlNode"/> tree.</summary>
    public static YamlNode SerializeToNode<T>(T value, YamlTypeInfo<T> typeInfo) {
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        var builder = new YamlNodeBuilder();
        using (var writer = new YamlWriter(builder)) Write(writer, value, typeInfo);
        return builder.Root!;
    }

    /// <summary>Reads a YAML document with generated metadata. Empty input reads as null.</summary>
    /// <exception cref="YamlException">The input is malformed, exceeds a limit or does not fit the type.</exception>
    public static T? Deserialize<T>(string yaml, YamlTypeInfo<T> typeInfo) {
        if (yaml is null) throw new ArgumentNullException(nameof(yaml));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        using var reader = new YamlReader(yaml, typeInfo.Options.EffectiveReaderOptions);
        return ReadDocument(reader, typeInfo);
    }

    /// <summary>Reads a YAML document from a stream with generated metadata; the encoding is detected. The stream stays open.</summary>
    public static T? Deserialize<T>(Stream utf8Yaml, YamlTypeInfo<T> typeInfo) {
        if (utf8Yaml is null) throw new ArgumentNullException(nameof(utf8Yaml));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        using var reader = new YamlReader(utf8Yaml, typeInfo.Options.EffectiveReaderOptions);
        return ReadDocument(reader, typeInfo);
    }

    /// <summary>Reads a YAML document from a text reader with generated metadata. The reader stays open.</summary>
    public static T? Deserialize<T>(TextReader reader, YamlTypeInfo<T> typeInfo) {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        using var yaml = new YamlReader(reader, typeInfo.Options.EffectiveReaderOptions);
        return ReadDocument(yaml, typeInfo);
    }

    /// <summary>Reads one value from a <see cref="YamlReader"/> with generated metadata, as <see cref="Deserialize{T}(YamlReader, YamlSerializerOptions?)"/> does.</summary>
    public static T? Deserialize<T>(YamlReader reader, YamlTypeInfo<T> typeInfo) {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        return ReadNext(reader, typeInfo);
    }

    /// <summary>Reads a model node with generated metadata.</summary>
    public static T? Deserialize<T>(YamlNode node, YamlTypeInfo<T> typeInfo) {
        if (node is null) throw new ArgumentNullException(nameof(node));
        if (typeInfo is null) throw new ArgumentNullException(nameof(typeInfo));
        int maxExpansion = typeInfo.Options.EffectiveReaderOptions.MaxAliasExpansion;
        var cursor = new YamlEventCursor(YamlEventCursor.Record(node), maxExpansion);
        return typeInfo.Read(cursor, new YamlReadState(typeInfo.Options));
    }

    // ------------------------------------------------------------------ generated metadata: context and Type

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> with the context's metadata and returns the text.</summary>
    public static string Serialize(object? value, Type inputType, YamlSerializerContext context) => SerializeBoxed(value, GetTypeInfo(inputType, context, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> with the context's metadata to a UTF-8 stream, which stays open.</summary>
    public static void Serialize(Stream utf8Yaml, object? value, Type inputType, YamlSerializerContext context) => SerializeBoxed(utf8Yaml, value, GetTypeInfo(inputType, context, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> with the context's metadata to a text writer, which stays open.</summary>
    public static void Serialize(TextWriter writer, object? value, Type inputType, YamlSerializerContext context) => SerializeBoxed(writer, value, GetTypeInfo(inputType, context, value));

    /// <summary>Writes <paramref name="value"/> as <paramref name="inputType"/> with the context's metadata, as one node at the writer's position.</summary>
    public static void Serialize(YamlWriter writer, object? value, Type inputType, YamlSerializerContext context) => SerializeBoxed(writer, value, GetTypeInfo(inputType, context, value));

    /// <summary>Converts <paramref name="value"/> as <paramref name="inputType"/> with the context's metadata to a detached <see cref="YamlNode"/> tree.</summary>
    public static YamlNode SerializeToNode(object? value, Type inputType, YamlSerializerContext context) => SerializeBoxedToNode(value, GetTypeInfo(inputType, context, value));

    /// <summary>Reads a YAML document as <paramref name="returnType"/> with the context's metadata.</summary>
    public static object? Deserialize(string yaml, Type returnType, YamlSerializerContext context) => DeserializeBoxed(yaml, GetTypeInfo(returnType, context, null));

    /// <summary>Reads a YAML document from a stream as <paramref name="returnType"/> with the context's metadata. The stream stays open.</summary>
    public static object? Deserialize(Stream utf8Yaml, Type returnType, YamlSerializerContext context) => DeserializeBoxed(utf8Yaml, GetTypeInfo(returnType, context, null));

    /// <summary>Reads a YAML document from a text reader as <paramref name="returnType"/> with the context's metadata. The reader stays open.</summary>
    public static object? Deserialize(TextReader reader, Type returnType, YamlSerializerContext context) => DeserializeBoxed(reader, GetTypeInfo(returnType, context, null));

    /// <summary>Reads one value from a <see cref="YamlReader"/> as <paramref name="returnType"/> with the context's metadata.</summary>
    public static object? Deserialize(YamlReader reader, Type returnType, YamlSerializerContext context) => DeserializeBoxed(reader, GetTypeInfo(returnType, context, null));

    /// <summary>Reads a model node as <paramref name="returnType"/> with the context's metadata.</summary>
    public static object? Deserialize(YamlNode node, Type returnType, YamlSerializerContext context) => DeserializeBoxed(node, GetTypeInfo(returnType, context, null));

    // ------------------------------------------------------------------ cores

    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    private static YamlTypeInfo<T> GetTypeInfo<T>(YamlSerializerOptions? options) => (YamlTypeInfo<T>)GetTypeInfo(typeof(T), options, null);

    [RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
    [RequiresDynamicCode(YamlReflectionResolver.DynamicCodeMessage)]
    private static YamlTypeInfo GetTypeInfo(Type type, YamlSerializerOptions? options, object? value) {
        if (type is null) throw new ArgumentNullException(nameof(type));
        CheckValue(type, value);
        options ??= YamlSerializerOptions.Default;
        if (options.TypeInfoResolver is null && IsReflectionEnabledByDefault) options.UseReflectionFallback(YamlReflectionResolver.GetInstance());
        return options.GetTypeInfo(type);
    }

    private static YamlTypeInfo GetTypeInfo(Type type, YamlSerializerContext context, object? value) {
        if (type is null) throw new ArgumentNullException(nameof(type));
        if (context is null) throw new ArgumentNullException(nameof(context));
        CheckValue(type, value);
        return context.Options.GetTypeInfo(type);
    }

    private static void CheckValue(Type type, object? value) {
        if (value is not null && !type.IsInstanceOfType(value)) {
            throw new ArgumentException("The value is a " + YamlSerializationErrors.TypeName(value.GetType()) + ", not a " + YamlSerializationErrors.TypeName(type) + ".", nameof(value));
        }
    }

    private static void Write<T>(YamlWriter writer, T value, YamlTypeInfo<T> typeInfo) {
        var state = new YamlWriteState(typeInfo.Options);
        if (state.Preserve) {
            var collector = new YamlReferenceCollector(typeInfo.Options);
            typeInfo.Collect(value, collector);
            state.SetShared(collector.Shared);
        }
        typeInfo.Write(writer, value, state);
    }

    private static void WriteBoxed(YamlWriter writer, object? value, YamlTypeInfo typeInfo) {
        var state = new YamlWriteState(typeInfo.Options);
        if (state.Preserve) {
            var collector = new YamlReferenceCollector(typeInfo.Options);
            typeInfo.CollectBoxed(value, collector);
            state.SetShared(collector.Shared);
        }
        typeInfo.WriteBoxed(writer, value, state);
    }

    private static string SerializeBoxed(object? value, YamlTypeInfo typeInfo) {
        using var output = new YamlCharBuffer();
        using (var writer = new YamlWriter(output, typeInfo.Options.EffectiveWriterOptions)) WriteBoxed(writer, value, typeInfo);
        return output.ToString();
    }

    private static void SerializeBoxed(Stream utf8Yaml, object? value, YamlTypeInfo typeInfo) {
        if (utf8Yaml is null) throw new ArgumentNullException(nameof(utf8Yaml));
        using var writer = new YamlWriter(utf8Yaml, typeInfo.Options.EffectiveWriterOptions);
        WriteBoxed(writer, value, typeInfo);
    }

    private static void SerializeBoxed(TextWriter writer, object? value, YamlTypeInfo typeInfo) {
        if (writer is null) throw new ArgumentNullException(nameof(writer));
        using var yaml = new YamlWriter(writer, typeInfo.Options.EffectiveWriterOptions);
        WriteBoxed(yaml, value, typeInfo);
    }

    private static void SerializeBoxed(YamlWriter writer, object? value, YamlTypeInfo typeInfo) {
        if (writer is null) throw new ArgumentNullException(nameof(writer));
        WriteBoxed(writer, value, typeInfo);
    }

    private static YamlNode SerializeBoxedToNode(object? value, YamlTypeInfo typeInfo) {
        var builder = new YamlNodeBuilder();
        using (var writer = new YamlWriter(builder)) WriteBoxed(writer, value, typeInfo);
        return builder.Root!;
    }

    private static object? DeserializeBoxed(string yaml, YamlTypeInfo typeInfo) {
        if (yaml is null) throw new ArgumentNullException(nameof(yaml));
        using var reader = new YamlReader(yaml, typeInfo.Options.EffectiveReaderOptions);
        return ReadDocumentBoxed(reader, typeInfo);
    }

    private static object? DeserializeBoxed(Stream utf8Yaml, YamlTypeInfo typeInfo) {
        if (utf8Yaml is null) throw new ArgumentNullException(nameof(utf8Yaml));
        using var reader = new YamlReader(utf8Yaml, typeInfo.Options.EffectiveReaderOptions);
        return ReadDocumentBoxed(reader, typeInfo);
    }

    private static object? DeserializeBoxed(TextReader reader, YamlTypeInfo typeInfo) {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        using var yaml = new YamlReader(reader, typeInfo.Options.EffectiveReaderOptions);
        return ReadDocumentBoxed(yaml, typeInfo);
    }

    private static object? DeserializeBoxed(YamlReader reader, YamlTypeInfo typeInfo) {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        if (!MoveToValue(reader, out bool document)) return null;
        object? value = typeInfo.ReadBoxed(Begin(reader), new YamlReadState(typeInfo.Options));
        if (document) ReadDocumentEnd(reader);
        return value;
    }

    private static object? DeserializeBoxed(YamlNode node, YamlTypeInfo typeInfo) {
        if (node is null) throw new ArgumentNullException(nameof(node));
        int maxExpansion = typeInfo.Options.EffectiveReaderOptions.MaxAliasExpansion;
        var cursor = new YamlEventCursor(YamlEventCursor.Record(node), maxExpansion);
        return typeInfo.ReadBoxed(cursor, new YamlReadState(typeInfo.Options));
    }

    private static T? ReadDocument<T>(YamlReader reader, YamlTypeInfo<T> typeInfo) {
        if (!MoveToValue(reader, out _)) return ReadEmpty(typeInfo);
        T value = typeInfo.Read(Begin(reader), new YamlReadState(typeInfo.Options));
        ReadDocumentEnd(reader);
        RequireStreamEnd(reader);
        return value;
    }

    private static object? ReadDocumentBoxed(YamlReader reader, YamlTypeInfo typeInfo) {
        if (!MoveToValue(reader, out _)) return null;
        object? value = typeInfo.ReadBoxed(Begin(reader), new YamlReadState(typeInfo.Options));
        ReadDocumentEnd(reader);
        RequireStreamEnd(reader);
        return value;
    }

    private static T? ReadNext<T>(YamlReader reader, YamlTypeInfo<T> typeInfo) {
        if (!MoveToValue(reader, out bool document)) return ReadEmpty(typeInfo);
        T value = typeInfo.Read(Begin(reader), new YamlReadState(typeInfo.Options));
        if (document) ReadDocumentEnd(reader);
        return value;
    }

    /// <summary>An empty stream reads as null, which a non-nullable value type cannot hold.</summary>
    private static T? ReadEmpty<T>(YamlTypeInfo<T> typeInfo) {
        if (default(T) is not null) {
            throw new YamlException(YamlErrorCodes.TypeMismatch, "The YAML input has no document to read as " + YamlSerializationErrors.TypeName(typeof(T)) + ".", default);
        }
        return default;
    }

    private static YamlEventCursor Begin(YamlReader reader) {
        var cursor = new YamlEventCursor(reader);
        cursor.Begin();
        return cursor;
    }

    /// <summary>
    /// Moves a reader to the node to read. Between documents it advances to the next document's root and reports
    /// <paramref name="document"/>; it returns false at the end of the stream.
    /// </summary>
    private static bool MoveToValue(YamlReader reader, out bool document) {
        document = false;
        while (true) {
            switch (reader.TokenType) {
                case YamlTokenType.None:
                case YamlTokenType.StreamStart:
                case YamlTokenType.DocumentEnd:
                case YamlTokenType.Comment:
                    if (!reader.Read()) return false;
                    continue;
                case YamlTokenType.DocumentStart:
                    document = true;
                    if (!reader.Read()) return false;
                    continue;
                case YamlTokenType.StreamEnd:
                    return false;
                case YamlTokenType.Scalar:
                case YamlTokenType.Alias:
                case YamlTokenType.MappingStart:
                case YamlTokenType.SequenceStart:
                    return true;
                default:
                    throw new InvalidOperationException("The reader is on a " + reader.TokenType + " event; position it between documents or on the node to read.");
            }
        }
    }

    private static void ReadDocumentEnd(YamlReader reader) {
        while (reader.Read() && reader.TokenType == YamlTokenType.Comment) {
        }
    }

    private static void RequireStreamEnd(YamlReader reader) {
        while (reader.Read()) {
            if (reader.TokenType == YamlTokenType.DocumentStart) {
                throw new YamlException(YamlErrorCodes.InvalidDocument,
                    "The YAML input holds more than one document; read a multi-document stream with a YamlReader, one value per document.", reader.Start);
            }
        }
    }
}
