using System.Net.Http.Headers;
using System.Text;

namespace OfficeIMO.Confluence;

/// <summary>Applies caller-owned authentication to an outgoing Confluence request.</summary>
public interface IConfluenceCredentialSource {
    /// <summary>Applies credentials to a request before it is sent.</summary>
    /// <remarks>The client calls this method again for each retry of a read request.</remarks>
    Task ApplyAsync(HttpRequestMessage request, CancellationToken cancellationToken = default);
}

/// <summary>Applies Atlassian email/API-token basic authentication.</summary>
public sealed class ConfluenceBasicCredentialSource : IConfluenceCredentialSource {
    private readonly string _encodedCredential;

    /// <summary>Creates a Basic authentication source from an Atlassian account email and API token.</summary>
    public ConfluenceBasicCredentialSource(string email, string apiToken) {
        if (string.IsNullOrWhiteSpace(email)) throw new ArgumentException("Atlassian account email is required.", nameof(email));
        if (string.IsNullOrWhiteSpace(apiToken)) throw new ArgumentException("Atlassian API token is required.", nameof(apiToken));
        _encodedCredential = Convert.ToBase64String(Encoding.UTF8.GetBytes(email + ":" + apiToken));
    }

    /// <summary>Sets the request's Basic Authorization header.</summary>
    public Task ApplyAsync(HttpRequestMessage request, CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        request.Headers.Authorization = new AuthenticationHeaderValue("Basic", _encodedCredential);
        return Task.CompletedTask;
    }
}

/// <summary>Applies a caller-supplied OAuth bearer token.</summary>
public sealed class ConfluenceBearerCredentialSource : IConfluenceCredentialSource {
    private readonly string _accessToken;

    /// <summary>Creates a Bearer authentication source from the supplied access token.</summary>
    public ConfluenceBearerCredentialSource(string accessToken) {
        if (string.IsNullOrWhiteSpace(accessToken)) throw new ArgumentException("Access token is required.", nameof(accessToken));
        _accessToken = accessToken;
    }

    /// <summary>Sets the request's Bearer Authorization header.</summary>
    public Task ApplyAsync(HttpRequestMessage request, CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
        return Task.CompletedTask;
    }
}

/// <summary>Delegates authentication to a caller-owned token or signing workflow.</summary>
public sealed class ConfluenceDelegateCredentialSource : IConfluenceCredentialSource {
    private readonly Func<HttpRequestMessage, CancellationToken, Task> _apply;

    /// <summary>Creates a credential source that invokes <paramref name="apply"/> for each outgoing request.</summary>
    public ConfluenceDelegateCredentialSource(Func<HttpRequestMessage, CancellationToken, Task> apply) =>
        _apply = apply ?? throw new ArgumentNullException(nameof(apply));

    /// <summary>Invokes the caller-supplied credential callback for this request.</summary>
    public Task ApplyAsync(HttpRequestMessage request, CancellationToken cancellationToken = default) => _apply(request, cancellationToken);
}
