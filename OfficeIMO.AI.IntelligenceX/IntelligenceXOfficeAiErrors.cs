using System.Net;
using IntelligenceX.OpenAI;

namespace OfficeIMO.AI.IntelligenceX;

/// <summary>Maps public SDK and HTTP failure contracts to content-free document-AI diagnostics.</summary>
public static class IntelligenceXOfficeAiErrors {
    /// <summary>
    /// Classifies only typed evidence. It never parses exception messages, provider payloads or credential URLs.
    /// Internal or untyped SDK failures remain unknown rather than guessing from potentially sensitive text.
    /// </summary>
    public static OfficeAiExecutionFailure Classify(Exception exception) {
        ArgumentNullException.ThrowIfNull(exception);
        return exception switch {
            OfficeAiExecutionException failure => failure.Failure,
            OpenAIAuthenticationRequiredException => OfficeAiExecutionFailure.AuthenticationRequired,
            HttpRequestException { StatusCode: HttpStatusCode.Unauthorized } => OfficeAiExecutionFailure.AuthenticationRequired,
            HttpRequestException { StatusCode: HttpStatusCode.Forbidden } => OfficeAiExecutionFailure.AccessDenied,
            HttpRequestException { StatusCode: HttpStatusCode.NotFound } => OfficeAiExecutionFailure.NotFound,
            HttpRequestException { StatusCode: HttpStatusCode.TooManyRequests } => OfficeAiExecutionFailure.RateLimited,
            HttpRequestException { StatusCode: HttpStatusCode.RequestTimeout or HttpStatusCode.GatewayTimeout } => OfficeAiExecutionFailure.TimedOut,
            HttpRequestException { StatusCode: HttpStatusCode.BadRequest or HttpStatusCode.UnprocessableEntity } => OfficeAiExecutionFailure.InvalidRequest,
            HttpRequestException { StatusCode: null or HttpStatusCode.BadGateway or HttpStatusCode.ServiceUnavailable or HttpStatusCode.InternalServerError }
                => OfficeAiExecutionFailure.Unavailable,
            TimeoutException or OperationCanceledException => OfficeAiExecutionFailure.TimedOut,
            _ => OfficeAiExecutionFailure.Unknown
        };
    }
}
