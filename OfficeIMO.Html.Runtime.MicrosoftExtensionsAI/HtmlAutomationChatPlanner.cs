using System.Collections;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.AI;

namespace OfficeIMO.Html.Runtime.MicrosoftExtensionsAI;

/// <summary>Caller-owned messages and model options for one automation planning turn.</summary>
public sealed class HtmlAutomationChatRequest {
    /// <summary>Messages chosen by the consuming application.</summary>
    public IReadOnlyList<ChatMessage> Messages { get; init; } = Array.Empty<ChatMessage>();
    /// <summary>Model, sampling and provider options chosen by the consuming application.</summary>
    public ChatOptions? Options { get; init; }
}

/// <summary>Creates a model request from an OfficeIMO automation turn.</summary>
public delegate HtmlAutomationChatRequest HtmlAutomationChatRequestFactory(HtmlAutomationTurn turn);

/// <summary>Bounds model tool-call conversion before calls reach the runtime runner.</summary>
public sealed class HtmlAutomationChatPlannerOptions {
    /// <summary>Maximum function calls accepted from one model response.</summary>
    public int MaxToolCallsPerTurn { get; set; } = 8;
    /// <summary>Maximum UTF-8 bytes accepted for one function-call argument object.</summary>
    public int MaxArgumentBytes { get; set; } = 64 * 1024;
    /// <summary>Maximum nested object and array depth accepted in function-call arguments.</summary>
    public int MaxArgumentDepth { get; set; } = 16;
    /// <summary>Maximum total object properties and array items accepted in one argument object.</summary>
    public int MaxArgumentItems { get; set; } = 4096;

    internal HtmlAutomationChatPlannerOptions Snapshot() {
        if (MaxToolCallsPerTurn <= 0 || MaxToolCallsPerTurn > 256) throw new ArgumentOutOfRangeException(nameof(MaxToolCallsPerTurn));
        if (MaxArgumentBytes <= 0 || MaxArgumentBytes > 16 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(MaxArgumentBytes));
        if (MaxArgumentDepth <= 0 || MaxArgumentDepth > 64) throw new ArgumentOutOfRangeException(nameof(MaxArgumentDepth));
        if (MaxArgumentItems <= 0 || MaxArgumentItems > 100_000) throw new ArgumentOutOfRangeException(nameof(MaxArgumentItems));
        return new HtmlAutomationChatPlannerOptions {
            MaxToolCallsPerTurn = MaxToolCallsPerTurn, MaxArgumentBytes = MaxArgumentBytes,
            MaxArgumentDepth = MaxArgumentDepth, MaxArgumentItems = MaxArgumentItems
        };
    }
}

/// <summary>Adapts any Microsoft.Extensions.AI chat client to the bounded OfficeIMO planner callback.</summary>
public sealed class HtmlAutomationChatPlanner {
    private readonly IChatClient _client;
    private readonly HtmlAutomationChatRequestFactory _requests;
    private readonly IReadOnlyList<AITool> _tools;
    private readonly HtmlAutomationChatPlannerOptions _limits;

    /// <summary>Creates an adapter without taking ownership of the caller's chat client.</summary>
    public HtmlAutomationChatPlanner(IChatClient client, HtmlAutomationChatRequestFactory requests,
        HtmlAutomationChatPlannerOptions? options = null) {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _requests = requests ?? throw new ArgumentNullException(nameof(requests));
        _limits = (options ?? new HtmlAutomationChatPlannerOptions()).Snapshot();
        _tools = Array.AsReadOnly(HtmlAutomationToolCatalog.GetDefinitions().Select(item => (AITool)
            AIFunctionFactory.CreateDeclaration(item.Name, item.Description, item.InputSchema)).ToArray());
    }

    /// <summary>Declaration-only model tools using the exact OfficeIMO JSON Schemas.</summary>
    public IReadOnlyList<AITool> Tools => _tools;

    /// <summary>Returns the callback consumed by <see cref="HtmlAutomationRunner"/>.</summary>
    public HtmlAutomationPlanner CreatePlanner() => PlanAsync;

    private async Task<HtmlAutomationPlannerDecision> PlanAsync(HtmlAutomationTurn turn, CancellationToken cancellationToken) {
        HtmlAutomationChatRequest request = _requests(turn) ?? throw new InvalidOperationException("The application returned no model request.");
        if (request.Messages == null || request.Messages.Count == 0)
            throw new InvalidOperationException("The application must supply at least one chat message.");
        ChatOptions options = request.Options?.Clone() ?? new ChatOptions();
        options.Tools = options.Tools?.ToList() ?? new List<AITool>();
        foreach (AITool tool in _tools) {
            string name = ((AIFunctionDeclaration)tool).Name;
            if (options.Tools.OfType<AIFunctionDeclaration>().Any(item => item.Name == name))
                throw new InvalidOperationException($"The application supplied a conflicting reserved OfficeIMO tool named '{name}'.");
            options.Tools.Add(tool);
        }
        ChatResponse response = await _client.GetResponseAsync(request.Messages, options, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException("The model client returned no response.");
        FunctionCallContent[] calls = response.Messages.SelectMany(message => message.Contents)
            .OfType<FunctionCallContent>().Where(item => !item.InformationalOnly).ToArray();
        if (calls.Length == 0) return HtmlAutomationPlannerDecision.Complete(response.Text);
        if (calls.Length > _limits.MaxToolCallsPerTurn)
            throw new InvalidOperationException("The model response exceeds the configured tool-call limit.");
        FunctionCallContent? invalid = calls.FirstOrDefault(item => item.Exception != null);
        if (invalid != null)
            throw new InvalidOperationException($"The model returned an invalid OfficeIMO tool call '{invalid.Name}'.", invalid.Exception);
        return HtmlAutomationPlannerDecision.Execute(calls.Select(item => new HtmlAutomationToolCall(
            item.CallId, item.Name, ToJson(item.Arguments, _limits))).ToArray());
    }

    private static JsonElement ToJson(IDictionary<string, object?>? arguments, HtmlAutomationChatPlannerOptions limits) {
        using var stream = new MemoryStream();
        var state = new ArgumentWriteState(stream, limits);
        using (var writer = new Utf8JsonWriter(stream)) {
            WriteDictionary(writer, arguments ?? new Dictionary<string, object?>(), state, 1);
            writer.Flush();
        }
        state.CheckBytes();
        using JsonDocument document = JsonDocument.Parse(stream.ToArray());
        return document.RootElement.Clone();
    }

    private static void WriteDictionary(Utf8JsonWriter writer, IDictionary<string, object?> values,
        ArgumentWriteState state, int depth) {
        state.CheckDepth(depth);
        if (values.Count > state.RemainingItems)
            throw new InvalidOperationException("The model tool arguments exceed the configured item limit.");
        writer.WriteStartObject();
        foreach (KeyValuePair<string, object?> item in values) {
            state.AddItem();
            state.CheckString(item.Key);
            writer.WritePropertyName(item.Key);
            WriteValue(writer, item.Value, state, depth + 1);
        }
        writer.WriteEndObject();
        state.CheckBytes(writer);
    }

    private static void WriteValue(Utf8JsonWriter writer, object? value, ArgumentWriteState state, int depth) {
        state.CheckDepth(depth);
        switch (value) {
            case null: writer.WriteNullValue(); break;
            case JsonElement json: WriteJson(writer, json, state, depth); break;
            case string text: state.CheckString(text); writer.WriteStringValue(text); break;
            case bool boolean: writer.WriteBooleanValue(boolean); break;
            case int number: writer.WriteNumberValue(number); break;
            case long number: writer.WriteNumberValue(number); break;
            case double number when double.IsFinite(number): writer.WriteNumberValue(number); break;
            case float number when float.IsFinite(number): writer.WriteNumberValue(number); break;
            case decimal number: writer.WriteNumberValue(number); break;
            case IDictionary<string, object?> dictionary: WriteDictionary(writer, dictionary, state, depth); break;
            case IEnumerable sequence:
                writer.WriteStartArray();
                foreach (object? item in sequence) { state.AddItem(); WriteValue(writer, item, state, depth + 1); }
                writer.WriteEndArray();
                break;
            default: throw new InvalidOperationException($"Unsupported model tool argument type '{value.GetType().FullName}'.");
        }
        state.CheckBytes(writer);
    }

    private static void WriteJson(Utf8JsonWriter writer, JsonElement value, ArgumentWriteState state, int depth) {
        switch (value.ValueKind) {
            case JsonValueKind.Object:
                state.CheckDepth(depth);
                writer.WriteStartObject();
                foreach (JsonProperty property in value.EnumerateObject()) {
                    state.AddItem(); state.CheckString(property.Name); writer.WritePropertyName(property.Name);
                    WriteJson(writer, property.Value, state, depth + 1);
                }
                writer.WriteEndObject();
                break;
            case JsonValueKind.Array:
                state.CheckDepth(depth);
                writer.WriteStartArray();
                foreach (JsonElement item in value.EnumerateArray()) { state.AddItem(); WriteJson(writer, item, state, depth + 1); }
                writer.WriteEndArray();
                break;
            case JsonValueKind.String:
                string text = value.GetString() ?? string.Empty; state.CheckString(text); writer.WriteStringValue(text); break;
            case JsonValueKind.Number: value.WriteTo(writer); break;
            case JsonValueKind.True: writer.WriteBooleanValue(true); break;
            case JsonValueKind.False: writer.WriteBooleanValue(false); break;
            case JsonValueKind.Null: writer.WriteNullValue(); break;
            default: throw new InvalidOperationException("Undefined model tool arguments are not supported.");
        }
        state.CheckBytes(writer);
    }

    private sealed class ArgumentWriteState {
        private readonly MemoryStream _stream;
        private readonly HtmlAutomationChatPlannerOptions _limits;
        private int _items;

        internal ArgumentWriteState(MemoryStream stream, HtmlAutomationChatPlannerOptions limits) {
            _stream = stream;
            _limits = limits;
        }

        internal int RemainingItems => _limits.MaxArgumentItems - _items;

        internal void AddItem() {
            if (++_items > _limits.MaxArgumentItems)
                throw new InvalidOperationException("The model tool arguments exceed the configured item limit.");
        }
        internal void CheckDepth(int depth) {
            if (depth > _limits.MaxArgumentDepth)
                throw new InvalidOperationException("The model tool arguments exceed the configured depth limit.");
        }
        internal void CheckString(string value) {
            if (value.Length > _limits.MaxArgumentBytes || Encoding.UTF8.GetByteCount(value) > _limits.MaxArgumentBytes)
                throw new InvalidOperationException("The model tool arguments exceed the configured byte limit.");
        }
        internal void CheckBytes(Utf8JsonWriter? writer = null) {
            writer?.Flush();
            if (_stream.Length > _limits.MaxArgumentBytes)
                throw new InvalidOperationException("The model tool arguments exceed the configured byte limit.");
        }
    }
}
