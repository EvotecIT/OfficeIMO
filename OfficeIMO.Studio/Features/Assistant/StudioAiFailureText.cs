using OfficeIMO.AI;
using OfficeIMO.AI.IntelligenceX;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Assistant;

/// <summary>Localized next actions for safe provider categories; raw exception text never reaches the UI.</summary>
internal static class StudioAiFailureText {
    internal static string FromException(IStudioLocalizer localizer, Exception exception) =>
        FromCode(localizer, new OfficeAiExecutionException(IntelligenceXOfficeAiErrors.Classify(exception)).DiagnosticCode);

    internal static string FromCode(IStudioLocalizer localizer, string code) {
        (string key, string fallback) = code switch {
            "provider-authentication-required" => ("AuthenticationRequired", "Your login or credential was rejected. Open Connections and sign in again, or update the API key."),
            "provider-access-denied" => ("AccessDenied", "This account cannot access the requested service or model. Check its permissions or choose another account/model."),
            "provider-not-found" => ("NotFound", "The endpoint or model was not found. Check the base URL, reconnect to refresh the model list, and select an available model."),
            "provider-rate-limited" => ("RateLimited", "The provider reported a usage or rate limit. Wait before retrying or check the account's available usage. No automatic retry was made."),
            "provider-unavailable" => ("Unavailable", "The service could not be reached. Check the network, or start your local model server and verify its address, then reconnect."),
            "provider-timed-out" => ("TimedOut", "The service did not finish in time. Check its availability or try a smaller page scope."),
            "provider-invalid-request" => ("InvalidRequest", "The provider rejected this request. Verify the endpoint and model's document-response support before retrying."),
            _ => ("Unknown", "The provider could not complete the operation. Reconnect and check account and model support. No provider response or credentials are shown.")
        };
        return localizer.GetOrDefault("Assistant.Failure." + key, fallback);
    }
}
