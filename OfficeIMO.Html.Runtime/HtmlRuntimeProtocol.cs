using System.Buffers.Binary;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Html.Runtime;

// Length-prefixed UTF-8 frames avoid EOF-based capture and unbounded line buffering.
internal static class HtmlRuntimeProtocol {
    internal const int MaximumRequestCharacters = 64 * 1024 * 1024;
    private static readonly UTF8Encoding Utf8 = new(false, true);

    internal static async Task WriteAsync<T>(Stream stream, T message, int maximumCharacters, CancellationToken token) {
        string json = JsonSerializer.Serialize(message);
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
        return JsonSerializer.Deserialize<T>(json) ?? throw new HtmlScriptRuntimeException("An empty runtime message is invalid.");
    }
}

internal sealed class HtmlRuntimeCommand {
    public long Id { get; set; }
    public string Kind { get; set; } = string.Empty;
    public HtmlScriptRequest? Request { get; set; }
    public string? Script { get; set; }
    public HtmlAutomationRequest? Automation { get; set; }
    public bool ReplaceHistoryEntry { get; set; }
}

internal sealed class HtmlRuntimeResponse {
    public long Id { get; set; }
    public string? Error { get; set; }
    public string? ValueJson { get; set; }
    public HtmlRuntimeWireDocument? Document { get; set; }
    public HtmlAutomationResult? Automation { get; set; }
}
