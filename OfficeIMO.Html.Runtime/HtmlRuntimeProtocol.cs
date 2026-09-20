using System.Buffers.Binary;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization.Metadata;

namespace OfficeIMO.Html.Runtime;

// Length-prefixed UTF-8 frames avoid EOF-based capture and unbounded line buffering.
internal static class HtmlRuntimeProtocol {
    internal const int MaximumRequestCharacters = 64 * 1024 * 1024;
    private static readonly UTF8Encoding Utf8 = new(false, true);

    internal static async Task WriteAsync<T>(Stream stream, T message, int maximumCharacters, CancellationToken token) {
        string json = JsonSerializer.Serialize(message, TypeInfo<T>());
        if (json.Length > maximumCharacters) throw new HtmlScriptRuntimeException("The runtime message exceeds its character budget.");
        byte[] bytes = Utf8.GetBytes(json);
        byte[] header = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(header, bytes.Length);
        await stream.WriteAsync(header, token).ConfigureAwait(false);
        await stream.WriteAsync(bytes, token).ConfigureAwait(false);
        await stream.FlushAsync(token).ConfigureAwait(false);
    }

    internal static async Task<T?> ReadAsync<T>(Stream stream, int maximumCharacters, CancellationToken token) where T : class {
        byte[] header = new byte[4];
        int first = await stream.ReadAsync(header.AsMemory(0, 1), token).ConfigureAwait(false);
        if (first == 0) return null;
        await stream.ReadExactlyAsync(header.AsMemory(1), token).ConfigureAwait(false);
        int length = BinaryPrimitives.ReadInt32LittleEndian(header);
        if (length <= 0 || length > Math.Min(int.MaxValue, (long)maximumCharacters * 3))
            throw new HtmlScriptRuntimeException("The runtime frame exceeds its byte budget.");
        byte[] bytes = new byte[length];
        await stream.ReadExactlyAsync(bytes, token).ConfigureAwait(false);
        string json = Utf8.GetString(bytes);
        if (json.Length > maximumCharacters) throw new HtmlScriptRuntimeException("The runtime frame exceeds its character budget.");
        return JsonSerializer.Deserialize(json, TypeInfo<T>()) ?? throw new HtmlScriptRuntimeException("An empty runtime message is invalid.");
    }

    internal static int MeasureCharacters<T>(T message) => JsonSerializer.Serialize(message, TypeInfo<T>()).Length;

    private static JsonTypeInfo<T> TypeInfo<T>() =>
        (JsonTypeInfo<T>)(HtmlRuntimeProtocolJsonContext.Default.GetTypeInfo(typeof(T))
            ?? throw new HtmlScriptRuntimeException("The runtime protocol type is not registered."));
}

internal sealed class HtmlRuntimeCommand {
    public long Id { get; set; }
    public string Kind { get; set; } = string.Empty;
    public HtmlScriptRequest? Request { get; set; }
    public string? Script { get; set; }
    public HtmlAutomationRequest? Automation { get; set; }
    public HtmlPageObservationRequest? Observation { get; set; }
    public string? ContextId { get; set; }
    public string? PageId { get; set; }
    public bool ReplaceHistoryEntry { get; set; }
    public HtmlRuntimeWireTraceOptions? Trace { get; set; }
}

internal sealed class HtmlRuntimeResponse {
    public long Id { get; set; }
    public string? Error { get; set; }
    public string? ErrorKind { get; set; }
    public Uri[]? MissingResourceUrls { get; set; }
    public HtmlRuntimeFetchDiscovery[]? MissingFetchRequests { get; set; }
    public string[] ConsumedFetchReplayIdentities { get; set; } = Array.Empty<string>();
    public HtmlRuntimeNavigationDiscovery[]? MissingNavigationRequests { get; set; }
    public string[] ConsumedNavigationReplayIdentities { get; set; } = Array.Empty<string>();
    public string? ValueJson { get; set; }
    public HtmlRuntimeWireDocument? Document { get; set; }
    public HtmlAutomationResult? Automation { get; set; }
    public HtmlPageObservation? Observation { get; set; }
    public long PageRevision { get; set; }
    public List<HtmlRuntimeWireEvent> Events { get; set; } = new();
}

internal sealed class HtmlRuntimeWireTraceOptions {
    public bool Enabled { get; set; }
    public int MaxEvents { get; set; }
    public bool IncludeUrls { get; set; }
    public bool IncludeConsoleMessages { get; set; }
    public bool IncludeFailureMessages { get; set; }
    public int MaxDetailCharacters { get; set; }
}

internal sealed class HtmlRuntimeWireEvent {
    public HtmlRuntimeEventKind Kind { get; set; }
    public string Operation { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public DateTimeOffset StartedUtc { get; set; }
    public double ElapsedMilliseconds { get; set; }
    public long? PageRevision { get; set; }
    public string? Detail { get; set; }
    public Uri? Url { get; set; }
    public string? Method { get; set; }
    public int? StatusCode { get; set; }
    public long? ByteCount { get; set; }
    public int? RedirectCount { get; set; }
    public string? Decision { get; set; }
    public string? ArtifactId { get; set; }
}
