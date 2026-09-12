using System.Text.Json;

namespace OfficeIMO.Html.Runtime;

/// <summary>A live scripted document. Operations are serialized; captured documents are independent.</summary>
/// <remarks>Dispose the session when finished. Cancellation or failure after command admission terminates the session.</remarks>
public interface IHtmlRuntimeSession : IAsyncDisposable {
    /// <summary>Runs a classic script in the existing document without returning interpreter objects.</summary>
    Task ExecuteAsync(string script, CancellationToken cancellationToken = default);
    /// <summary>Evaluates an expression and returns a JSON value. Undefined and non-serializable results fail the operation.</summary>
    Task<JsonElement> EvaluateAsync(string expression, CancellationToken cancellationToken = default);
    /// <summary>Waits until an expression evaluates to boolean true, using the session's command deadline.</summary>
    Task WaitForAsync(string expression, CancellationToken cancellationToken = default);
    /// <summary>Waits for the supplied condition and captures a frozen document in the same event-loop task.</summary>
    /// <param name="readyExpression">A boolean condition; null uses the initial request's readiness expression.</param>
    /// <param name="cancellationToken">Cancellation while queued leaves the session usable; active cancellation terminates it.</param>
    Task<HtmlScriptCapture> CaptureAsync(string? readyExpression = null, CancellationToken cancellationToken = default);
}
