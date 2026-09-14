using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.AI.Html;

/// <summary>Creates application-owned instructions for one observed HTML automation turn.</summary>
public delegate string HtmlAutomationAiInstructionFactory(HtmlAutomationTurn turn);

/// <summary>Bounds the OfficeIMO AI bridge before a decision reaches the HTML runtime.</summary>
public sealed class HtmlAutomationAiPlannerOptions {
    /// <summary>Maximum function calls accepted from one model response.</summary>
    public int MaxToolCallsPerTurn { get; set; } = 8;

    /// <summary>Maximum characters accepted from one model response.</summary>
    public int MaxResponseCharacters { get; set; } = 64 * 1024;

    /// <summary>Maximum UTF-8 bytes accepted for one function-call argument object.</summary>
    public int MaxArgumentBytes { get; set; } = 64 * 1024;

    /// <summary>Maximum nested depth accepted in model JSON.</summary>
    public int MaxJsonDepth { get; set; } = 16;

    /// <summary>Maximum total properties and array items accepted in one argument object.</summary>
    public int MaxArgumentItems { get; set; } = 4096;

    internal HtmlAutomationAiPlannerOptions Snapshot() {
        if (MaxToolCallsPerTurn <= 0 || MaxToolCallsPerTurn > 256) throw new ArgumentOutOfRangeException(nameof(MaxToolCallsPerTurn));
        if (MaxResponseCharacters <= 0 || MaxResponseCharacters > 16 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(MaxResponseCharacters));
        if (MaxArgumentBytes <= 0 || MaxArgumentBytes > 16 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(MaxArgumentBytes));
        if (MaxJsonDepth <= 0 || MaxJsonDepth > 64) throw new ArgumentOutOfRangeException(nameof(MaxJsonDepth));
        if (MaxArgumentItems <= 0 || MaxArgumentItems > 100_000) throw new ArgumentOutOfRangeException(nameof(MaxArgumentItems));
        return new HtmlAutomationAiPlannerOptions {
            MaxToolCallsPerTurn = MaxToolCallsPerTurn,
            MaxResponseCharacters = MaxResponseCharacters,
            MaxArgumentBytes = MaxArgumentBytes,
            MaxJsonDepth = MaxJsonDepth,
            MaxArgumentItems = MaxArgumentItems
        };
    }
}

/// <summary>Adapts an OfficeIMO AI executor to the bounded HTML automation planner callback.</summary>
public sealed class HtmlAutomationAiPlanner {
    private readonly OfficeAiToolPlanner _planner;
    private readonly HtmlAutomationAiInstructionFactory _instructions;
    private readonly HtmlAutomationAiPlannerOptions _limits;
    private readonly IReadOnlyList<OfficeAiToolDefinition> _tools;

    /// <summary>Creates a bridge without taking ownership of the caller's executor.</summary>
    public HtmlAutomationAiPlanner(IOfficeAiExecutor executor, HtmlAutomationAiInstructionFactory instructions,
        HtmlAutomationAiPlannerOptions? options = null) {
        _planner = new OfficeAiToolPlanner(executor ?? throw new ArgumentNullException(nameof(executor)));
        _instructions = instructions ?? throw new ArgumentNullException(nameof(instructions));
        _limits = (options ?? new HtmlAutomationAiPlannerOptions()).Snapshot();
        _tools = Array.AsReadOnly(HtmlAutomationToolCatalog.GetDefinitions().Select(tool =>
            new OfficeAiToolDefinition(tool.Name, tool.Description, tool.InputSchema)).ToArray());
    }

    /// <summary>Owned tool declarations supplied to the OfficeIMO AI executor.</summary>
    public IReadOnlyList<OfficeAiToolDefinition> Tools => _tools;

    /// <summary>
    /// Returns the callback consumed by <see cref="HtmlAutomationRunner"/>. The runner's dispatcher validates each
    /// argument object against the selected HTML tool contract before invoking runtime behavior.
    /// </summary>
    public HtmlAutomationPlanner CreatePlanner() => PlanAsync;

    private async Task<HtmlAutomationPlannerDecision> PlanAsync(HtmlAutomationTurn turn, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(turn);
        string instructions = _instructions(turn);
        if (string.IsNullOrWhiteSpace(instructions)) throw new InvalidOperationException("The application returned no HTML automation instructions.");
        OfficeAiToolPlanningDecision decision = await _planner.PlanAsync(new OfficeAiToolPlanningRequest {
            RequestId = $"html-automation-{turn.Step}",
            Instructions = instructions,
            InputJson = SerializeTurn(turn),
            Tools = _tools,
            MaxToolCalls = _limits.MaxToolCallsPerTurn,
            MaxResponseCharacters = _limits.MaxResponseCharacters,
            MaxArgumentBytes = _limits.MaxArgumentBytes,
            MaxJsonDepth = _limits.MaxJsonDepth,
            MaxArgumentItems = _limits.MaxArgumentItems
        }, cancellationToken).ConfigureAwait(false);
        return decision.IsComplete
            ? HtmlAutomationPlannerDecision.Complete(decision.Message)
            : HtmlAutomationPlannerDecision.Execute(decision.Calls.Select(call =>
                new HtmlAutomationToolCall(call.Id, call.Name, call.Arguments)).ToArray());
    }

    private static string SerializeTurn(HtmlAutomationTurn turn) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream)) {
            writer.WriteStartObject();
            writer.WriteNumber("step", turn.Step);
            writer.WritePropertyName("observation");
            WriteJson(writer, HtmlRuntimeJson.Serialize(turn.Observation));
            writer.WritePropertyName("previousResults");
            writer.WriteStartArray();
            foreach (HtmlAutomationToolResult result in turn.PreviousResults) WriteJson(writer, HtmlRuntimeJson.Serialize(result));
            writer.WriteEndArray();
            writer.WritePropertyName("tools");
            writer.WriteStartArray();
            foreach (HtmlAutomationToolDefinition tool in turn.Tools) {
                writer.WriteStartObject();
                writer.WriteString("name", tool.Name);
                writer.WriteString("description", tool.Description);
                writer.WritePropertyName("inputSchema");
                tool.InputSchema.WriteTo(writer);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    private static void WriteJson(Utf8JsonWriter writer, string json) {
        using JsonDocument document = JsonDocument.Parse(json);
        document.RootElement.WriteTo(writer);
    }
}
