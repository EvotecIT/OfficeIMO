using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;

namespace OfficeIMO.Html.Runtime;

/// <summary>Stable names for optional provider-neutral automation tools.</summary>
public static class HtmlAutomationToolNames {
    /// <summary>Observe the current page.</summary>
    public const string ObservePage = "officeimo_observe_page";
    /// <summary>Run a structured action against a locator or observed reference.</summary>
    public const string Act = "officeimo_act";
    /// <summary>Navigate the current page.</summary>
    public const string Navigate = "officeimo_navigate";
    /// <summary>Capture the current document as an inert OfficeIMO snapshot.</summary>
    public const string Capture = "officeimo_capture";
}

/// <summary>A model-SDK-neutral function definition with a JSON Schema input contract.</summary>
public sealed class HtmlAutomationToolDefinition {
    /// <summary>Creates an immutable tool definition.</summary>
    public HtmlAutomationToolDefinition(string name, string description, string inputSchemaJson) {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentException.ThrowIfNullOrWhiteSpace(description);
        ArgumentException.ThrowIfNullOrWhiteSpace(inputSchemaJson);
        using JsonDocument schema = JsonDocument.Parse(inputSchemaJson);
        Name = name;
        Description = description;
        InputSchema = schema.RootElement.Clone();
    }

    /// <summary>Stable tool name.</summary>
    public string Name { get; }
    /// <summary>Human-readable operation description.</summary>
    public string Description { get; }
    /// <summary>JSON Schema for tool arguments.</summary>
    public JsonElement InputSchema { get; }
}

/// <summary>Built-in automation tool definitions supported by <see cref="HtmlAutomationToolDispatcher"/>.</summary>
public static class HtmlAutomationToolCatalog {
    private static readonly IReadOnlyList<HtmlAutomationToolDefinition> Tools = Array.AsReadOnly(new[] {
        new HtmlAutomationToolDefinition(HtmlAutomationToolNames.ObservePage, "Observe the current page as bounded semantic and visual data.",
            """{"type":"object","additionalProperties":false,"properties":{"mode":{"type":"string","enum":["Semantic","Visual","Combined"]},"maxElements":{"type":"integer","minimum":1,"maximum":10000},"maxTextCharacters":{"type":"integer","minimum":1},"includeHidden":{"type":"boolean"},"actionableOnly":{"type":"boolean"}}}"""),
        new HtmlAutomationToolDefinition(HtmlAutomationToolNames.Act, "Run one structured action against an observed element reference or locator.",
            """{"type":"object","additionalProperties":false,"required":["action"],"properties":{"reference":{"type":"object","additionalProperties":false,"required":["pageId","revision","elementIndex","elementName","elementId"],"properties":{"pageId":{"type":"string"},"revision":{"type":"integer","minimum":1},"elementIndex":{"type":"integer","minimum":0},"elementName":{"type":"string"},"elementId":{"type":"string"}}},"css":{"type":"string"},"action":{"type":"string","enum":["Inspect","Count","Click","Hover","Press","Fill","SetChecked","SelectOptions","Focus","Blur","ScrollIntoView","Wait","SetSelection"]},"value":{"type":"string"},"values":{"type":"array","items":{"type":"string"}},"checked":{"type":"boolean"},"modifiers":{"type":"string","enum":["None","Alt","Control","Meta","Shift","Alt, Control","Alt, Meta","Alt, Shift","Control, Meta","Control, Shift","Meta, Shift","Alt, Control, Meta","Alt, Control, Shift","Alt, Meta, Shift","Control, Meta, Shift","Alt, Control, Meta, Shift"]},"waitState":{"type":"string","enum":["Attached","Detached","Enabled","Disabled","Editable","Focused","Visible","Hidden","InViewport","Value","Text","Checked"]},"waitForReady":{"type":"boolean"},"selectionStart":{"type":"integer","minimum":0},"selectionEnd":{"type":"integer","minimum":0}}}"""),
        new HtmlAutomationToolDefinition(HtmlAutomationToolNames.Navigate, "Navigate the current WebApplicationV1 page to an allowed absolute URL.",
            """{"type":"object","additionalProperties":false,"required":["url"],"properties":{"url":{"type":"string","format":"uri"},"replaceHistoryEntry":{"type":"boolean"}}}"""),
        new HtmlAutomationToolDefinition(HtmlAutomationToolNames.Capture, "Capture the current live DOM as an inert OfficeIMO document snapshot.",
            """{"type":"object","additionalProperties":false,"properties":{"readyExpression":{"type":"string"}}}""")
    });

    /// <summary>Returns immutable definitions for every built-in tool.</summary>
    public static IReadOnlyList<HtmlAutomationToolDefinition> GetDefinitions() => Tools;
}

/// <summary>One planner or caller-issued tool call.</summary>
public sealed class HtmlAutomationToolCall {
    /// <summary>Creates a call from a stable name and JSON arguments.</summary>
    public HtmlAutomationToolCall(string id, string name, JsonElement arguments) {
        ArgumentException.ThrowIfNullOrWhiteSpace(id);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        Id = id;
        Name = name;
        Arguments = arguments.Clone();
    }

    /// <summary>Caller correlation identity.</summary>
    public string Id { get; }
    /// <summary>Tool name.</summary>
    public string Name { get; }
    /// <summary>Detached JSON arguments.</summary>
    public JsonElement Arguments { get; }

    /// <summary>Creates an observation call.</summary>
    public static HtmlAutomationToolCall Observe(string id, HtmlPageObservationRequest? request = null) =>
        Create(id, HtmlAutomationToolNames.ObservePage, request ?? new HtmlPageObservationRequest());

    /// <summary>Creates an action call.</summary>
    public static HtmlAutomationToolCall Act(string id, HtmlAutomationRequest request) =>
        Create(id, HtmlAutomationToolNames.Act, HtmlAutomationToolArguments.From(request ?? throw new ArgumentNullException(nameof(request))));

    /// <summary>Creates a navigation call.</summary>
    public static HtmlAutomationToolCall Navigate(string id, Uri url, bool replaceHistoryEntry = false) =>
        Create(id, HtmlAutomationToolNames.Navigate, new HtmlNavigationToolArguments { Url = url, ReplaceHistoryEntry = replaceHistoryEntry });

    /// <summary>Creates a document-capture call.</summary>
    public static HtmlAutomationToolCall Capture(string id, string? readyExpression = null) =>
        Create(id, HtmlAutomationToolNames.Capture, new HtmlCaptureToolArguments { ReadyExpression = readyExpression });

    private static HtmlAutomationToolCall Create<T>(string id, string name, T arguments) =>
        new(id, name, JsonSerializer.SerializeToElement(arguments, HtmlAutomationToolJson.TypeInfo<T>()));
}

/// <summary>Structured outcome of one automation tool call.</summary>
public sealed class HtmlAutomationToolResult {
    /// <summary>Call correlation identity.</summary>
    public string CallId { get; init; } = string.Empty;
    /// <summary>Tool name.</summary>
    public string ToolName { get; init; } = string.Empty;
    /// <summary>Whether dispatch completed without an exception.</summary>
    public bool IsSuccess { get; init; }
    /// <summary>Provider-neutral failure message.</summary>
    public string? Error { get; init; }
    /// <summary>Observation returned by an observation call.</summary>
    public HtmlPageObservation? Observation { get; init; }
    /// <summary>Automation outcome returned by an action call.</summary>
    public HtmlAutomationResult? Automation { get; init; }
    /// <summary>Inert snapshot returned by a capture call.</summary>
    public HtmlScriptCapture? Capture { get; init; }
}

/// <summary>Validates and dispatches built-in tools through <see cref="IHtmlRuntimePage"/>.</summary>
public sealed class HtmlAutomationToolDispatcher {
    /// <summary>Executes one tool call without exposing provider-native objects.</summary>
    public async Task<HtmlAutomationToolResult> ExecuteAsync(IHtmlRuntimePage page, HtmlAutomationToolCall call,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(page);
        ArgumentNullException.ThrowIfNull(call);
        try {
            return call.Name switch {
                HtmlAutomationToolNames.ObservePage => Success(call, observation: await page.ObserveAsync(
                    Deserialize<HtmlPageObservationRequest>(call), cancellationToken).ConfigureAwait(false)),
                HtmlAutomationToolNames.Act => Success(call, automation: await page.AutomateAsync(
                    Deserialize<HtmlAutomationToolArguments>(call).ToRequest(), cancellationToken).ConfigureAwait(false)),
                HtmlAutomationToolNames.Navigate => await NavigateAsync(page, call, cancellationToken).ConfigureAwait(false),
                HtmlAutomationToolNames.Capture => Success(call, capture: await page.CaptureAsync(
                    Deserialize<HtmlCaptureToolArguments>(call).ReadyExpression, cancellationToken).ConfigureAwait(false)),
                _ => throw new ArgumentException("Unknown OfficeIMO automation tool: " + call.Name, nameof(call))
            };
        } catch (Exception error) when (error is not OperationCanceledException) {
            return new HtmlAutomationToolResult { CallId = call.Id, ToolName = call.Name, Error = error.Message };
        }
    }

    private static async Task<HtmlAutomationToolResult> NavigateAsync(IHtmlRuntimePage page, HtmlAutomationToolCall call, CancellationToken token) {
        HtmlNavigationToolArguments arguments = Deserialize<HtmlNavigationToolArguments>(call);
        if (arguments.Url == null) throw new ArgumentException("Navigate requires an absolute URL.");
        await page.NavigateAsync(arguments.Url, arguments.ReplaceHistoryEntry, token).ConfigureAwait(false);
        return Success(call);
    }

    private static T Deserialize<T>(HtmlAutomationToolCall call) where T : class =>
        call.Arguments.Deserialize(HtmlAutomationToolJson.TypeInfo<T>())
        ?? throw new ArgumentException("The tool arguments are invalid.", nameof(call));

    private static HtmlAutomationToolResult Success(HtmlAutomationToolCall call, HtmlPageObservation? observation = null,
        HtmlAutomationResult? automation = null, HtmlScriptCapture? capture = null) => new() {
        CallId = call.Id, ToolName = call.Name, IsSuccess = true,
        Observation = observation, Automation = automation, Capture = capture
    };
}

internal static class HtmlAutomationToolJson {
    internal static JsonTypeInfo<T> TypeInfo<T>() =>
        (JsonTypeInfo<T>)(HtmlAutomationToolJsonContext.Default.GetTypeInfo(typeof(T))
            ?? throw new InvalidOperationException("The automation tool argument type is not registered."));
}

internal sealed class HtmlNavigationToolArguments {
    public Uri? Url { get; set; }
    public bool ReplaceHistoryEntry { get; set; }
}

internal sealed class HtmlCaptureToolArguments {
    public string? ReadyExpression { get; set; }
}

internal sealed class HtmlAutomationToolArguments {
    public HtmlObservedElementReference? Reference { get; set; }
    public string? Css { get; set; }
    public HtmlAutomationAction Action { get; set; }
    public string? Value { get; set; }
    public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();
    public bool? Checked { get; set; }
    public HtmlKeyboardModifiers Modifiers { get; set; }
    public HtmlLocatorWaitState WaitState { get; set; }
    public bool WaitForReady { get; set; } = true;
    public int? SelectionStart { get; set; }
    public int? SelectionEnd { get; set; }

    internal static HtmlAutomationToolArguments From(HtmlAutomationRequest request) => new() {
        Reference = request.Reference, Css = SimpleCss(request.Query),
        Action = request.Action, Value = request.Value, Values = request.Values, Checked = request.Checked, Modifiers = request.Modifiers,
        WaitState = request.WaitState, WaitForReady = request.WaitForReady,
        SelectionStart = request.SelectionStart, SelectionEnd = request.SelectionEnd
    };

    private static string? SimpleCss(HtmlLocatorQuery? query) {
        if (query == null) return null;
        if (query.Kind != HtmlLocatorKind.Css || query.Scope != null || query.Index != null)
            throw new NotSupportedException("The optional tool contract accepts observed references or an unscoped CSS locator. Use deterministic .NET locators for richer queries.");
        return query.Value;
    }

    internal HtmlAutomationRequest ToRequest() {
        if ((Reference == null) == string.IsNullOrWhiteSpace(Css))
            throw new ArgumentException("Act requires exactly one target: reference or css.");
        return new HtmlAutomationRequest {
            Reference = Reference,
            Query = Reference == null ? HtmlLocatorQuery.Css(Css!) : null,
            Action = Action,
            Value = Value,
            Values = Values ?? Array.Empty<string>(),
            Checked = Checked,
            Modifiers = Modifiers,
            WaitState = WaitState,
            WaitForReady = WaitForReady,
            SelectionStart = SelectionStart,
            SelectionEnd = SelectionEnd
        };
    }
}
