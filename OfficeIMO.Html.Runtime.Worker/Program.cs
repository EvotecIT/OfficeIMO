using System.Text.Json;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Worker;

// One request and one response per process. The parent owns deadline termination.
try {
    string json = await HtmlProcessRuntimeProvider.ReadBoundedAsync(Console.In, 64 * 1024 * 1024, CancellationToken.None);
    HtmlScriptRequest request = (JsonSerializer.Deserialize<HtmlScriptRequest>(json)
        ?? throw new HtmlScriptRuntimeException("A runtime request is required.")).Snapshot();
    HtmlRuntimeWireDocument capture = await ScriptedDocumentCapture.RunAsync(request);
    string response = JsonSerializer.Serialize(capture);
    if (response.Length > request.MaxOutputCharacters) throw new HtmlScriptRuntimeException("The capture exceeds MaxOutputCharacters.");
    await Console.Out.WriteAsync(response);
} catch (Exception error) {
    await Console.Out.WriteAsync(JsonSerializer.Serialize(new HtmlRuntimeWireDocument { Error = error.Message }));
}
