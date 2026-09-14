using System.Text.Json;

namespace OfficeIMO.Html.Runtime;

/// <summary>NativeAOT-safe JSON serialization for provider-neutral runtime outputs.</summary>
public static class HtmlRuntimeJson {
    /// <summary>Serializes a provider descriptor with camel-case properties and string enums.</summary>
    public static string Serialize(HtmlRuntimeProviderDescriptor value) => Serialize(value, HtmlRuntimePublicJsonContext.Default.HtmlRuntimeProviderDescriptor);
    /// <summary>Serializes a bounded page observation with camel-case properties and string enums.</summary>
    public static string Serialize(HtmlPageObservation value) => Serialize(value, HtmlRuntimePublicJsonContext.Default.HtmlPageObservation);
    /// <summary>Serializes a structured automation result with camel-case properties and string enums.</summary>
    public static string Serialize(HtmlAutomationResult value) => Serialize(value, HtmlRuntimePublicJsonContext.Default.HtmlAutomationResult);
    /// <summary>Serializes an immutable operation trace with camel-case properties and string enums.</summary>
    public static string Serialize(HtmlRuntimeTrace value) => Serialize(value, HtmlRuntimePublicJsonContext.Default.HtmlRuntimeTrace);
    /// <summary>Serializes an inert capture as HTML, resolved URLs, and retained resource responses.</summary>
    public static string Serialize(HtmlScriptCapture value) => Serialize(Project(value), HtmlRuntimePublicJsonContext.Default.HtmlScriptCapturePayload);
    /// <summary>Serializes a tool result, including an inert HTML projection when it contains a capture.</summary>
    public static string Serialize(HtmlAutomationToolResult value) => Serialize(Project(value), HtmlRuntimePublicJsonContext.Default.HtmlAutomationToolResultPayload);
    /// <summary>Serializes a completed planner run and its ordered tool results.</summary>
    public static string Serialize(HtmlAutomationRunResult value) => Serialize(new HtmlAutomationRunResultPayload {
        IsComplete = value?.IsComplete ?? throw new ArgumentNullException(nameof(value)),
        Steps = value.Steps,
        Message = value.Message,
        FinalObservation = value.FinalObservation,
        ToolResults = value.ToolResults.Select(Project).ToArray()
    }, HtmlRuntimePublicJsonContext.Default.HtmlAutomationRunResultPayload);

    private static string Serialize<T>(T value, System.Text.Json.Serialization.Metadata.JsonTypeInfo<T> typeInfo) {
        ArgumentNullException.ThrowIfNull(value);
        return JsonSerializer.Serialize(value, typeInfo);
    }

    private static HtmlAutomationToolResultPayload Project(HtmlAutomationToolResult value) {
        ArgumentNullException.ThrowIfNull(value);
        return new HtmlAutomationToolResultPayload {
            CallId = value.CallId,
            ToolName = value.ToolName,
            IsSuccess = value.IsSuccess,
            Error = value.Error,
            Observation = value.Observation,
            Automation = value.Automation,
            Capture = value.Capture == null ? null : Project(value.Capture)
        };
    }

    private static HtmlScriptCapturePayload Project(HtmlScriptCapture value) {
        ArgumentNullException.ThrowIfNull(value);
        return new HtmlScriptCapturePayload {
                DocumentHtml = value.Document.DocumentElement?.OuterHtml ?? string.Empty,
                ProviderId = value.ProviderId,
                DocumentUrl = value.DocumentUrl,
                BaseUri = value.BaseUri,
                Resources = value.Resources.Select(resource => new HtmlRuntimeResourcePayload {
                    Url = resource.Url,
                    FinalUrl = resource.FinalUrl,
                    RedirectCount = resource.RedirectCount,
                    Content = resource.Content,
                    ContentType = resource.ContentType,
                    StatusCode = resource.StatusCode,
                    Headers = resource.Headers,
                    StatusText = resource.StatusText
                }).ToArray()
        };
    }
}

internal sealed class HtmlAutomationToolResultPayload {
    public string CallId { get; set; } = string.Empty;
    public string ToolName { get; set; } = string.Empty;
    public bool IsSuccess { get; set; }
    public string? Error { get; set; }
    public HtmlPageObservation? Observation { get; set; }
    public HtmlAutomationResult? Automation { get; set; }
    public HtmlScriptCapturePayload? Capture { get; set; }
}

internal sealed class HtmlAutomationRunResultPayload {
    public bool IsComplete { get; set; }
    public int Steps { get; set; }
    public string? Message { get; set; }
    public HtmlPageObservation FinalObservation { get; set; } = null!;
    public IReadOnlyList<HtmlAutomationToolResultPayload> ToolResults { get; set; } = Array.Empty<HtmlAutomationToolResultPayload>();
}

internal sealed class HtmlScriptCapturePayload {
    public string DocumentHtml { get; set; } = string.Empty;
    public string ProviderId { get; set; } = string.Empty;
    public Uri DocumentUrl { get; set; } = new("https://officeimo.invalid/");
    public Uri BaseUri { get; set; } = new("https://officeimo.invalid/");
    public IReadOnlyList<HtmlRuntimeResourcePayload> Resources { get; set; } = Array.Empty<HtmlRuntimeResourcePayload>();
}

internal sealed class HtmlRuntimeResourcePayload {
    public Uri Url { get; set; } = new("https://officeimo.invalid/");
    public Uri FinalUrl { get; set; } = new("https://officeimo.invalid/");
    public int RedirectCount { get; set; }
    public byte[] Content { get; set; } = Array.Empty<byte>();
    public string ContentType { get; set; } = string.Empty;
    public int StatusCode { get; set; }
    public IReadOnlyDictionary<string, string> Headers { get; set; } = new Dictionary<string, string>();
    public string StatusText { get; set; } = string.Empty;
}
