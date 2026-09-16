using System.Text.Json.Serialization;

namespace OfficeIMO.Confluence;

/// <summary>Confluence page body representation.</summary>
public enum ConfluenceBodyFormat {
    /// <summary>Request the Confluence storage body representation.</summary>
    Storage,
    /// <summary>Request the Atlas Document Format body representation.</summary>
    AtlasDocFormat,
}

/// <summary>A Confluence page version.</summary>
public sealed class ConfluencePageVersion {
    /// <summary>Gets or sets the page or attachment version number.</summary>
    [JsonPropertyName("number")]
    public int Number { get; set; }
    /// <summary>Gets or sets the optional version message.</summary>
    [JsonPropertyName("message")]
    public string? Message { get; set; }
    /// <summary>Gets or sets whether the version is marked as a minor edit.</summary>
    [JsonPropertyName("minorEdit")]
    public bool MinorEdit { get; set; }
}

/// <summary>A represented Confluence page body.</summary>
public sealed class ConfluencePageBody {
    /// <summary>Gets or sets the body representation name, such as <c>storage</c> or <c>atlas_doc_format</c>.</summary>
    [JsonPropertyName("representation")]
    public string? Representation { get; set; }
    /// <summary>Gets or sets the body content in the specified representation.</summary>
    [JsonPropertyName("value")]
    public string? Value { get; set; }
}

/// <summary>Body representations returned for a page.</summary>
public sealed class ConfluencePageBodies {
    /// <summary>Gets or sets the storage-format body, when returned.</summary>
    [JsonPropertyName("storage")]
    public ConfluencePageBody? Storage { get; set; }
    /// <summary>Gets or sets the Atlas Document Format body, when returned.</summary>
    [JsonPropertyName("atlas_doc_format")]
    public ConfluencePageBody? AtlasDocFormat { get; set; }
    /// <summary>Gets or sets the rendered view body, when returned.</summary>
    [JsonPropertyName("view")]
    public ConfluencePageBody? View { get; set; }
}

/// <summary>A Confluence Cloud page.</summary>
public sealed class ConfluencePage {
    /// <summary>Gets or sets the Confluence page identifier.</summary>
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;
    /// <summary>Gets or sets the page status returned by Confluence.</summary>
    [JsonPropertyName("status")]
    public string Status { get; set; } = string.Empty;
    /// <summary>Gets or sets the page title.</summary>
    [JsonPropertyName("title")]
    public string Title { get; set; } = string.Empty;
    /// <summary>Gets or sets the containing space identifier.</summary>
    [JsonPropertyName("spaceId")]
    public string SpaceId { get; set; } = string.Empty;
    /// <summary>Gets or sets the parent page identifier, when supplied.</summary>
    [JsonPropertyName("parentId")]
    public string? ParentId { get; set; }
    /// <summary>Gets or sets the page version metadata.</summary>
    [JsonPropertyName("version")]
    public ConfluencePageVersion Version { get; set; } = new ConfluencePageVersion();
    /// <summary>Gets or sets the body representations returned for this page.</summary>
    [JsonPropertyName("body")]
    public ConfluencePageBodies Body { get; set; } = new ConfluencePageBodies();
}

/// <summary>A cursor-addressable page batch.</summary>
public sealed class ConfluencePageBatch {
    internal ConfluencePageBatch(IReadOnlyList<ConfluencePage> pages, string? nextRelativeUri) {
        Pages = pages;
        NextRelativeUri = nextRelativeUri;
    }
    /// <summary>Gets the pages returned in this batch.</summary>
    public IReadOnlyList<ConfluencePage> Pages { get; }
    /// <summary>Gets the next-page link supplied by Confluence, or <see langword="null"/> at the end.</summary>
    /// <remarks>Despite the property name, an HTTP Link header can supply an absolute URI.</remarks>
    public string? NextRelativeUri { get; }
    /// <summary>Decoded cursor for requesting the next batch, or null when enumeration is complete.</summary>
    public string? NextCursor => ConfluenceCursor.Extract(NextRelativeUri);
}

internal static class ConfluenceCursor {
    public static string? Extract(string? relativeUri) {
        if (string.IsNullOrWhiteSpace(relativeUri)) return null;
        int question = relativeUri!.IndexOf('?');
        if (question < 0 || question == relativeUri.Length - 1) return null;
        foreach (string item in relativeUri.Substring(question + 1).Split('&')) {
            string[] parts = item.Split(new[] { '=' }, 2);
            if (parts.Length == 2 && string.Equals(Uri.UnescapeDataString(parts[0]), "cursor", StringComparison.OrdinalIgnoreCase)) {
                string value = Uri.UnescapeDataString(parts[1]);
                return value.Length == 0 ? null : value;
            }
        }
        return null;
    }
}

internal static class ConfluencePagination {
    internal static string? Next(string? responseLink, IReadOnlyDictionary<string, IReadOnlyList<string>> headers) {
        if (!string.IsNullOrWhiteSpace(responseLink)) return responseLink;
        if (!headers.TryGetValue("Link", out IReadOnlyList<string>? values)) return null;
        foreach (string value in values) {
            foreach (string segment in value.Split(',')) {
                string candidate = segment.Trim();
                int close = candidate.IndexOf('>');
                if (!candidate.StartsWith("<", StringComparison.Ordinal) || close <= 1) continue;
                string parameters = candidate.Substring(close + 1);
                bool isNext = parameters.Split(';')
                    .Select(item => item.Trim())
                    .Any(item => string.Equals(item, "rel=next", StringComparison.OrdinalIgnoreCase) ||
                                 string.Equals(item, "rel=\"next\"", StringComparison.OrdinalIgnoreCase));
                if (isNext) return candidate.Substring(1, close - 1);
            }
        }
        return null;
    }
}

/// <summary>Options for listing Confluence pages.</summary>
public sealed class ConfluencePageQuery {
    /// <summary>Gets or sets an optional space identifier filter.</summary>
    public string? SpaceId { get; set; }
    /// <summary>Gets or sets an optional title filter sent to Confluence.</summary>
    public string? Title { get; set; }
    /// <summary>Gets or sets the cursor value to send with the next listing request.</summary>
    public string? Cursor { get; set; }
    /// <summary>Gets or sets the batch limit, from 1 through 250; the default is 25.</summary>
    public int Limit { get; set; } = 25;
    /// <summary>Gets or sets the requested body representation; the default is Atlas Document Format.</summary>
    public ConfluenceBodyFormat BodyFormat { get; set; } = ConfluenceBodyFormat.AtlasDocFormat;
}

/// <summary>Input for creating a Confluence page.</summary>
public sealed class ConfluencePageCreateRequest {
    /// <summary>Gets or sets the required destination space identifier.</summary>
    public string SpaceId { get; set; } = string.Empty;
    /// <summary>Gets or sets the required page title.</summary>
    public string Title { get; set; } = string.Empty;
    /// <summary>Gets or sets an optional parent page identifier; a blank value is omitted.</summary>
    public string? ParentId { get; set; }
    /// <summary>Gets or sets the page status; a blank value is sent as <c>current</c>.</summary>
    public string Status { get; set; } = "current";
    /// <summary>Gets or sets the required body with a representation and non-null value.</summary>
    public ConfluencePageBody Body { get; set; } = new ConfluencePageBody();
}

/// <summary>Input for updating a Confluence page.</summary>
public sealed class ConfluencePageUpdateRequest {
    /// <summary>Gets or sets the required identifier of the page to update.</summary>
    public string PageId { get; set; } = string.Empty;
    /// <summary>Gets or sets the required new page title.</summary>
    public string Title { get; set; } = string.Empty;
    /// <summary>Gets or sets the page status; a blank value is sent as <c>current</c>.</summary>
    public string Status { get; set; } = "current";
    /// <summary>Gets or sets the caller-selected next positive version number; the client does not fetch it.</summary>
    public int VersionNumber { get; set; }
    /// <summary>Gets or sets the optional message included with the new version.</summary>
    public string? VersionMessage { get; set; }
    /// <summary>Gets or sets the required body with a representation and non-null value.</summary>
    public ConfluencePageBody Body { get; set; } = new ConfluencePageBody();
}

/// <summary>A serializable, non-executing representation of a pending page write.</summary>
public sealed class ConfluencePageWritePlan {
    internal ConfluencePageWritePlan(string method, string relativeUri, string payload) {
        Method = method;
        RelativeUri = relativeUri;
        Payload = payload;
    }
    /// <summary>Gets the HTTP method that would be used for the page write.</summary>
    public string Method { get; }
    /// <summary>Gets the API path and query to use relative to the session API base.</summary>
    public string RelativeUri { get; }
    /// <summary>Gets the JSON request body, or an empty string for a delete plan.</summary>
    public string Payload { get; }
}
