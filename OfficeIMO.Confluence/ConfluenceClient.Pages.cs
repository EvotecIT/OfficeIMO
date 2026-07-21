using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Confluence;

/// <summary>Dependency-light Confluence Cloud page and attachment client.</summary>
public sealed partial class ConfluenceClient : IDisposable {
    private readonly ConfluenceHttpTransport _transport;

    public ConfluenceClient(ConfluenceSession session) {
        Session = session ?? throw new ArgumentNullException(nameof(session));
        _transport = new ConfluenceHttpTransport(session);
    }

    public ConfluenceSession Session { get; }

    /// <summary>Reads a page with the requested body representation.</summary>
    public Task<ConfluencePage> GetPageAsync(string pageId, ConfluenceBodyFormat bodyFormat = ConfluenceBodyFormat.AtlasDocFormat, CancellationToken cancellationToken = default) {
        ValidateId(pageId, nameof(pageId));
        return _transport.SendJsonAsync<ConfluencePage>(HttpMethod.Get, "/wiki/api/v2/pages/" + Encode(pageId) + "?body-format=" + Format(bodyFormat), null, ConfluenceRequestSafety.SafeToRetry, cancellationToken);
    }

    /// <summary>Lists one cursor-addressable batch of pages.</summary>
    public async Task<ConfluencePageBatch> GetPagesAsync(ConfluencePageQuery? query = null, CancellationToken cancellationToken = default) {
        query ??= new ConfluencePageQuery();
        if (query.Limit < 1 || query.Limit > 250) throw new ArgumentOutOfRangeException(nameof(query), "Limit must be between 1 and 250.");
        string relativeUri = BuildPageQuery(query);
        CollectionResponse<ConfluencePage> response = await _transport.SendJsonAsync<CollectionResponse<ConfluencePage>>(HttpMethod.Get, relativeUri, null, ConfluenceRequestSafety.SafeToRetry, cancellationToken).ConfigureAwait(false);
        return new ConfluencePageBatch(response.Results, response.Links?.Next);
    }

    /// <summary>Creates a page. Non-idempotent writes are never retried automatically.</summary>
    public Task<ConfluencePage> CreatePageAsync(ConfluencePageCreateRequest request, CancellationToken cancellationToken = default) {
        ConfluencePageWritePlan plan = PlanCreatePage(request);
        using JsonDocument payload = JsonDocument.Parse(plan.Payload);
        return _transport.SendJsonAsync<ConfluencePage>(HttpMethod.Post, plan.RelativeUri, payload.RootElement.Clone(), ConfluenceRequestSafety.NonIdempotent, cancellationToken);
    }

    /// <summary>Updates a page using the caller-supplied next version number. Writes are never retried automatically.</summary>
    public Task<ConfluencePage> UpdatePageAsync(ConfluencePageUpdateRequest request, CancellationToken cancellationToken = default) {
        ConfluencePageWritePlan plan = PlanUpdatePage(request);
        using JsonDocument payload = JsonDocument.Parse(plan.Payload);
        return _transport.SendJsonAsync<ConfluencePage>(HttpMethod.Put, plan.RelativeUri, payload.RootElement.Clone(), ConfluenceRequestSafety.NonIdempotent, cancellationToken);
    }

    /// <summary>Builds the exact create request without contacting Confluence.</summary>
    public static ConfluencePageWritePlan PlanCreatePage(ConfluencePageCreateRequest request) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        ValidateId(request.SpaceId, nameof(request.SpaceId));
        ValidateTitle(request.Title);
        ValidateBody(request.Body);
        var payload = new {
            spaceId = request.SpaceId,
            status = string.IsNullOrWhiteSpace(request.Status) ? "current" : request.Status,
            title = request.Title,
            parentId = string.IsNullOrWhiteSpace(request.ParentId) ? null : request.ParentId,
            body = request.Body,
        };
        return new ConfluencePageWritePlan("POST", "/wiki/api/v2/pages", JsonSerializer.Serialize(payload, ConfluenceHttpTransport.JsonOptions));
    }

    /// <summary>Builds the exact update request without contacting Confluence.</summary>
    public static ConfluencePageWritePlan PlanUpdatePage(ConfluencePageUpdateRequest request) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        ValidateId(request.PageId, nameof(request.PageId));
        ValidateTitle(request.Title);
        ValidateBody(request.Body);
        if (request.VersionNumber < 1) throw new ArgumentOutOfRangeException(nameof(request.VersionNumber), "Version number must be the next positive page version.");
        var payload = new {
            id = request.PageId,
            status = string.IsNullOrWhiteSpace(request.Status) ? "current" : request.Status,
            title = request.Title,
            body = request.Body,
            version = new { number = request.VersionNumber, message = request.VersionMessage },
        };
        return new ConfluencePageWritePlan("PUT", "/wiki/api/v2/pages/" + Encode(request.PageId), JsonSerializer.Serialize(payload, ConfluenceHttpTransport.JsonOptions));
    }

    public void Dispose() => _transport.Dispose();

    private static string BuildPageQuery(ConfluencePageQuery query) {
        var values = new List<string> {
            "limit=" + query.Limit.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "body-format=" + Format(query.BodyFormat),
        };
        if (!string.IsNullOrWhiteSpace(query.SpaceId)) values.Add("space-id=" + Encode(query.SpaceId!));
        if (!string.IsNullOrWhiteSpace(query.Title)) values.Add("title=" + Encode(query.Title!));
        if (!string.IsNullOrWhiteSpace(query.Cursor)) values.Add("cursor=" + Encode(query.Cursor!));
        return "/wiki/api/v2/pages?" + string.Join("&", values);
    }

    private static string Format(ConfluenceBodyFormat format) => format == ConfluenceBodyFormat.Storage ? "storage" : "atlas_doc_format";
    private static string Encode(string value) => Uri.EscapeDataString(value);
    private static void ValidateId(string value, string parameterName) { if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Confluence identifier is required.", parameterName); }
    private static void ValidateTitle(string value) { if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Page title is required.", nameof(value)); }
    private static void ValidateBody(ConfluencePageBody body) {
        if (body == null) throw new ArgumentNullException(nameof(body));
        if (string.IsNullOrWhiteSpace(body.Representation)) throw new ArgumentException("Page body representation is required.", nameof(body));
        if (body.Value == null) throw new ArgumentException("Page body value is required.", nameof(body));
    }

    private sealed class CollectionResponse<T> {
        [JsonPropertyName("results")]
        public List<T> Results { get; set; } = new List<T>();
        [JsonPropertyName("_links")]
        public CollectionLinks? Links { get; set; }
    }

    private sealed class CollectionLinks {
        [JsonPropertyName("next")]
        public string? Next { get; set; }
    }
}
