using System.Collections.ObjectModel;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.AI;

/// <summary>Provider-neutral tool declaration for a bounded AI planning request.</summary>
public sealed class OfficeAiToolDefinition {
    /// <summary>Creates a tool declaration with a JSON Schema argument contract.</summary>
    public OfficeAiToolDefinition(string name, string description, JsonElement inputSchema) {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentException.ThrowIfNullOrWhiteSpace(description);
        if (name.Length > 128 || name.Any(char.IsControl)) throw new ArgumentException("Tool names must contain at most 128 non-control characters.", nameof(name));
        if (description.Length > 4096 || description.Any(char.IsControl)) throw new ArgumentException("Tool descriptions must contain at most 4096 non-control characters.", nameof(description));
        if (inputSchema.ValueKind != JsonValueKind.Object) throw new ArgumentException("A tool input schema must be a JSON object.", nameof(inputSchema));
        OfficeAiToolSchema.ValidateDefinition(inputSchema);
        Name = name;
        Description = description;
        InputSchema = inputSchema.Clone();
    }

    /// <summary>Stable tool name.</summary>
    public string Name { get; }

    /// <summary>Human-readable tool purpose.</summary>
    public string Description { get; }

    /// <summary>Closed JSON Schema from the dependency-free OfficeIMO tool-schema subset.</summary>
    public JsonElement InputSchema { get; }
}

/// <summary>One provider-neutral, bounded tool-planning request.</summary>
public sealed class OfficeAiToolPlanningRequest {
    /// <summary>Stable non-secret request identifier.</summary>
    public required string RequestId { get; init; }

    /// <summary>Application-owned instructions for this planning turn.</summary>
    public required string Instructions { get; init; }

    /// <summary>Current application state serialized as JSON.</summary>
    public required string InputJson { get; init; }

    /// <summary>Tools that the returned decision may call.</summary>
    public IReadOnlyList<OfficeAiToolDefinition> Tools { get; init; } = Array.Empty<OfficeAiToolDefinition>();

    /// <summary>Maximum tool calls accepted from one response.</summary>
    public int MaxToolCalls { get; init; } = 8;

    /// <summary>Maximum characters accepted from one provider response.</summary>
    public int MaxResponseCharacters { get; init; } = 64 * 1024;

    /// <summary>Maximum UTF-8 bytes accepted for one tool argument object.</summary>
    public int MaxArgumentBytes { get; init; } = 64 * 1024;

    /// <summary>Maximum nested depth accepted in provider JSON.</summary>
    public int MaxJsonDepth { get; init; } = 16;

    /// <summary>Maximum properties and array items accepted in one argument object.</summary>
    public int MaxArgumentItems { get; init; } = 4096;
}

/// <summary>One bounded, declaration-matched tool call returned by an AI executor.</summary>
public sealed class OfficeAiToolCall {
    internal OfficeAiToolCall(string id, string name, JsonElement arguments) {
        Id = id;
        Name = name;
        Arguments = arguments.Clone();
    }

    /// <summary>Correlation identity within the decision.</summary>
    public string Id { get; }

    /// <summary>Declared tool name.</summary>
    public string Name { get; }

    /// <summary>Detached JSON argument object.</summary>
    public JsonElement Arguments { get; }
}

/// <summary>A structurally validated completion or tool-call decision returned by an AI executor.</summary>
public sealed class OfficeAiToolPlanningDecision {
    internal OfficeAiToolPlanningDecision(bool isComplete, string? message, IReadOnlyList<OfficeAiToolCall> calls) {
        IsComplete = isComplete;
        Message = message;
        Calls = calls;
    }

    /// <summary>Whether the planner declared the application goal complete.</summary>
    public bool IsComplete { get; }

    /// <summary>Optional final message.</summary>
    public string? Message { get; }

    /// <summary>
    /// Structurally validated calls for the owning application to validate against its tool contract before execution.
    /// </summary>
    public IReadOnlyList<OfficeAiToolCall> Calls { get; }
}

/// <summary>Uses an OfficeIMO AI executor to produce locally bounded, structurally validated tool decisions.</summary>
public sealed class OfficeAiToolPlanner {
    private readonly IOfficeAiExecutor _executor;

    /// <summary>Creates a planner over a caller-owned OfficeIMO AI executor.</summary>
    public OfficeAiToolPlanner(IOfficeAiExecutor executor) {
        _executor = executor ?? throw new ArgumentNullException(nameof(executor));
        _executor.Profile.Validate();
    }

    /// <summary>Requests and validates one bounded decision.</summary>
    public async Task<OfficeAiToolPlanningDecision> PlanAsync(OfficeAiToolPlanningRequest request,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        Snapshot snapshot = Validate(request);
        string outputSchema = CreateOutputSchema(snapshot.Tools);
        var execution = new OfficeAiExecutionRequest(snapshot.RequestId, snapshot.Instructions, snapshot.InputJson,
            outputSchema, Array.Empty<OfficeAiImage>(), snapshot.MaxResponseCharacters);
        int measured = _executor.MeasureRequestCharacters(execution);
        if (measured < 0 || measured > _executor.Profile.MaxRequestCharacters)
            throw new InvalidOperationException("The tool-planning request exceeds the executor profile request limit.");

        OfficeAiExecutionResponse response = await _executor.ExecuteAsync(execution, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException("The AI executor returned no response.");
        if (!response.IsComplete) throw new InvalidOperationException("The AI executor returned a truncated tool-planning response.");
        if (response.Json.Length == 0 || response.Json.Length > snapshot.MaxResponseCharacters)
            throw new InvalidOperationException("The AI executor response exceeds the tool-planning response limit.");
        return ParseDecision(response.Json, snapshot);
    }

    private static Snapshot Validate(OfficeAiToolPlanningRequest request) {
        if (string.IsNullOrWhiteSpace(request.RequestId) || request.RequestId.Length > 256 || request.RequestId.Any(char.IsControl))
            throw new ArgumentException("RequestId must contain 1-256 non-control characters.", nameof(request));
        if (string.IsNullOrWhiteSpace(request.Instructions) || request.Instructions.Length > 1_000_000)
            throw new ArgumentException("Instructions must contain 1-1,000,000 characters.", nameof(request));
        if (string.IsNullOrWhiteSpace(request.InputJson) || request.InputJson.Length > 2_000_000)
            throw new ArgumentException("InputJson must contain 1-2,000,000 characters.", nameof(request));
        if (request.Tools == null || request.Tools.Count == 0 || request.Tools.Count > 256)
            throw new ArgumentException("A tool-planning request must declare 1-256 tools.", nameof(request));
        if (request.Tools.Any(tool => tool == null) || request.Tools.Select(tool => tool.Name).Distinct(StringComparer.Ordinal).Count() != request.Tools.Count)
            throw new ArgumentException("Tool declarations must be non-null and uniquely named.", nameof(request));
        if (request.MaxToolCalls <= 0 || request.MaxToolCalls > 256) throw new ArgumentOutOfRangeException(nameof(request.MaxToolCalls));
        if (request.MaxResponseCharacters <= 0 || request.MaxResponseCharacters > 16 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(request.MaxResponseCharacters));
        if (request.MaxArgumentBytes <= 0 || request.MaxArgumentBytes > 16 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(request.MaxArgumentBytes));
        if (request.MaxJsonDepth <= 0 || request.MaxJsonDepth > 64) throw new ArgumentOutOfRangeException(nameof(request.MaxJsonDepth));
        if (request.MaxArgumentItems <= 0 || request.MaxArgumentItems > 100_000) throw new ArgumentOutOfRangeException(nameof(request.MaxArgumentItems));
        using JsonDocument input = JsonDocument.Parse(request.InputJson, new JsonDocumentOptions { MaxDepth = request.MaxJsonDepth });
        if (input.RootElement.ValueKind is not (JsonValueKind.Object or JsonValueKind.Array))
            throw new ArgumentException("InputJson must contain a JSON object or array.", nameof(request));
        return new Snapshot(request.RequestId, request.Instructions, request.InputJson,
            Array.AsReadOnly(request.Tools.ToArray()), request.MaxToolCalls, request.MaxResponseCharacters,
            request.MaxArgumentBytes, request.MaxJsonDepth, request.MaxArgumentItems);
    }

    private static OfficeAiToolPlanningDecision ParseDecision(string json, Snapshot snapshot) {
        using JsonDocument document = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = snapshot.MaxJsonDepth });
        JsonElement root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object) throw new InvalidOperationException("The tool-planning response must be a JSON object.");
        RejectUnknownOrDuplicateProperties(root, "isComplete", "message", "calls");
        if (!root.TryGetProperty("isComplete", out JsonElement completed) || completed.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw new InvalidOperationException("The tool-planning response requires a boolean isComplete value.");
        string? message = null;
        if (root.TryGetProperty("message", out JsonElement messageElement)) {
            if (messageElement.ValueKind == JsonValueKind.String) message = messageElement.GetString();
            else if (messageElement.ValueKind != JsonValueKind.Null) throw new InvalidOperationException("The tool-planning message must be a string or null.");
            if (message?.Length > snapshot.MaxResponseCharacters) throw new InvalidOperationException("The tool-planning message exceeds the response limit.");
        }
        if (!root.TryGetProperty("calls", out JsonElement callsElement) || callsElement.ValueKind != JsonValueKind.Array)
            throw new InvalidOperationException("The tool-planning response requires a calls array.");
        int count = callsElement.GetArrayLength();
        if (count > snapshot.MaxToolCalls) throw new InvalidOperationException("The tool-planning response exceeds the configured call limit.");
        if (completed.GetBoolean() == (count != 0))
            throw new InvalidOperationException("A completed decision must have no calls and a continuing decision must have at least one call.");

        var tools = snapshot.Tools.ToDictionary(tool => tool.Name, StringComparer.Ordinal);
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var calls = new List<OfficeAiToolCall>(count);
        foreach (JsonElement item in callsElement.EnumerateArray()) {
            if (item.ValueKind != JsonValueKind.Object) throw new InvalidOperationException("Every tool call must be a JSON object.");
            RejectUnknownOrDuplicateProperties(item, "id", "name", "arguments");
            string id = RequiredString(item, "id", 256);
            string name = RequiredString(item, "name", 128);
            if (!ids.Add(id)) throw new InvalidOperationException("Tool call identifiers must be unique within a decision.");
            if (!tools.TryGetValue(name, out OfficeAiToolDefinition? tool)) throw new InvalidOperationException($"The AI executor returned an undeclared tool call '{name}'.");
            if (!item.TryGetProperty("arguments", out JsonElement arguments) || arguments.ValueKind != JsonValueKind.Object)
                throw new InvalidOperationException("Every tool call requires an arguments object.");
            if (Encoding.UTF8.GetByteCount(arguments.GetRawText()) > snapshot.MaxArgumentBytes)
                throw new InvalidOperationException("A tool argument object exceeds the configured byte limit.");
            if (CountItems(arguments, snapshot.MaxArgumentItems) > snapshot.MaxArgumentItems)
                throw new InvalidOperationException("A tool argument object exceeds the configured item limit.");
            JsonElement normalizedArguments = OfficeAiToolSchema.NormalizeAndValidate(arguments, tool.InputSchema);
            calls.Add(new OfficeAiToolCall(id, name, normalizedArguments));
        }
        return new OfficeAiToolPlanningDecision(completed.GetBoolean(), message,
            new ReadOnlyCollection<OfficeAiToolCall>(calls));
    }

    private static void RejectUnknownOrDuplicateProperties(JsonElement value, params string[] allowed) {
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonProperty property in value.EnumerateObject()) {
            if (!names.Add(property.Name)) throw new InvalidOperationException($"The tool-planning response contains duplicate '{property.Name}' properties.");
            if (!allowed.Contains(property.Name, StringComparer.Ordinal)) throw new InvalidOperationException($"The tool-planning response contains unknown '{property.Name}' data.");
        }
    }

    private static string RequiredString(JsonElement value, string name, int maximumLength) {
        if (!value.TryGetProperty(name, out JsonElement property) || property.ValueKind != JsonValueKind.String)
            throw new InvalidOperationException($"A tool call requires a string {name} value.");
        string result = property.GetString()!;
        if (string.IsNullOrWhiteSpace(result) || result.Length > maximumLength || result.Any(char.IsControl))
            throw new InvalidOperationException($"A tool call has an invalid {name} value.");
        return result;
    }

    private static int CountItems(JsonElement root, int limit) {
        int count = 0;
        var pending = new Stack<JsonElement>();
        pending.Push(root);
        while (pending.Count > 0) {
            JsonElement current = pending.Pop();
            if (current.ValueKind == JsonValueKind.Object) {
                foreach (JsonProperty property in current.EnumerateObject()) {
                    if (++count > limit) return count;
                    pending.Push(property.Value);
                }
            } else if (current.ValueKind == JsonValueKind.Array) {
                foreach (JsonElement item in current.EnumerateArray()) {
                    if (++count > limit) return count;
                    pending.Push(item);
                }
            }
        }
        return count;
    }

    private static string CreateOutputSchema(IReadOnlyList<OfficeAiToolDefinition> tools) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream)) {
            writer.WriteStartObject();
            writer.WriteString("type", "object");
            writer.WriteBoolean("additionalProperties", false);
            writer.WritePropertyName("required"); writer.WriteStartArray(); writer.WriteStringValue("isComplete"); writer.WriteStringValue("message"); writer.WriteStringValue("calls"); writer.WriteEndArray();
            writer.WritePropertyName("properties"); writer.WriteStartObject();
            writer.WritePropertyName("isComplete"); writer.WriteStartObject(); writer.WriteString("type", "boolean"); writer.WriteEndObject();
            writer.WritePropertyName("message"); writer.WriteStartObject(); writer.WritePropertyName("type"); writer.WriteStartArray(); writer.WriteStringValue("string"); writer.WriteStringValue("null"); writer.WriteEndArray(); writer.WriteEndObject();
            writer.WritePropertyName("calls"); writer.WriteStartObject();
            writer.WriteString("type", "array");
            writer.WritePropertyName("items"); writer.WriteStartObject(); writer.WritePropertyName("anyOf"); writer.WriteStartArray();
            foreach (OfficeAiToolDefinition tool in tools) {
                writer.WriteStartObject(); writer.WriteString("type", "object"); writer.WriteBoolean("additionalProperties", false);
                writer.WritePropertyName("required"); writer.WriteStartArray(); writer.WriteStringValue("id"); writer.WriteStringValue("name"); writer.WriteStringValue("arguments"); writer.WriteEndArray();
                writer.WritePropertyName("properties"); writer.WriteStartObject();
                writer.WritePropertyName("id"); writer.WriteStartObject(); writer.WriteString("type", "string"); writer.WriteEndObject();
                writer.WritePropertyName("name"); writer.WriteStartObject(); writer.WriteString("const", tool.Name); writer.WriteEndObject();
                writer.WritePropertyName("arguments"); OfficeAiToolSchema.WriteStrict(writer, tool.InputSchema, nullable: false);
                writer.WriteEndObject(); writer.WriteEndObject();
            }
            writer.WriteEndArray(); writer.WriteEndObject(); writer.WriteEndObject();
            writer.WriteEndObject(); writer.WriteEndObject();
        }
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    private sealed record Snapshot(string RequestId, string Instructions, string InputJson,
        IReadOnlyList<OfficeAiToolDefinition> Tools, int MaxToolCalls, int MaxResponseCharacters,
        int MaxArgumentBytes, int MaxJsonDepth, int MaxArgumentItems);
}
