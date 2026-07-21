using System.Text.Json.Serialization;

namespace OfficeIMO.Confluence;

/// <summary>Confluence page body representation.</summary>
public enum ConfluenceBodyFormat {
    Storage,
    AtlasDocFormat,
}

/// <summary>A Confluence page version.</summary>
public sealed class ConfluencePageVersion {
    [JsonPropertyName("number")]
    public int Number { get; set; }
    [JsonPropertyName("message")]
    public string? Message { get; set; }
    [JsonPropertyName("minorEdit")]
    public bool MinorEdit { get; set; }
}

/// <summary>A represented Confluence page body.</summary>
public sealed class ConfluencePageBody {
    [JsonPropertyName("representation")]
    public string? Representation { get; set; }
    [JsonPropertyName("value")]
    public string? Value { get; set; }
}

/// <summary>Body representations returned for a page.</summary>
public sealed class ConfluencePageBodies {
    [JsonPropertyName("storage")]
    public ConfluencePageBody? Storage { get; set; }
    [JsonPropertyName("atlas_doc_format")]
    public ConfluencePageBody? AtlasDocFormat { get; set; }
    [JsonPropertyName("view")]
    public ConfluencePageBody? View { get; set; }
}

/// <summary>A Confluence Cloud page.</summary>
public sealed class ConfluencePage {
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;
    [JsonPropertyName("status")]
    public string Status { get; set; } = string.Empty;
    [JsonPropertyName("title")]
    public string Title { get; set; } = string.Empty;
    [JsonPropertyName("spaceId")]
    public string SpaceId { get; set; } = string.Empty;
    [JsonPropertyName("parentId")]
    public string? ParentId { get; set; }
    [JsonPropertyName("version")]
    public ConfluencePageVersion Version { get; set; } = new ConfluencePageVersion();
    [JsonPropertyName("body")]
    public ConfluencePageBodies Body { get; set; } = new ConfluencePageBodies();
}

/// <summary>A cursor-addressable page batch.</summary>
public sealed class ConfluencePageBatch {
    internal ConfluencePageBatch(IReadOnlyList<ConfluencePage> pages, string? nextRelativeUri) {
        Pages = pages;
        NextRelativeUri = nextRelativeUri;
    }
    public IReadOnlyList<ConfluencePage> Pages { get; }
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

/// <summary>Options for listing Confluence pages.</summary>
public sealed class ConfluencePageQuery {
    public string? SpaceId { get; set; }
    public string? Title { get; set; }
    public string? Cursor { get; set; }
    public int Limit { get; set; } = 25;
    public ConfluenceBodyFormat BodyFormat { get; set; } = ConfluenceBodyFormat.AtlasDocFormat;
}

/// <summary>Input for creating a Confluence page.</summary>
public sealed class ConfluencePageCreateRequest {
    public string SpaceId { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string? ParentId { get; set; }
    public string Status { get; set; } = "current";
    public ConfluencePageBody Body { get; set; } = new ConfluencePageBody();
}

/// <summary>Input for updating a Confluence page.</summary>
public sealed class ConfluencePageUpdateRequest {
    public string PageId { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Status { get; set; } = "current";
    public int VersionNumber { get; set; }
    public string? VersionMessage { get; set; }
    public ConfluencePageBody Body { get; set; } = new ConfluencePageBody();
}

/// <summary>A serializable, non-executing representation of a pending page write.</summary>
public sealed class ConfluencePageWritePlan {
    internal ConfluencePageWritePlan(string method, string relativeUri, string payload) {
        Method = method;
        RelativeUri = relativeUri;
        Payload = payload;
    }
    public string Method { get; }
    public string RelativeUri { get; }
    public string Payload { get; }
}
