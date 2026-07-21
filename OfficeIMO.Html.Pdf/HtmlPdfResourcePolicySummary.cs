using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Html.Pdf;

/// <summary>Immutable snapshot of the shared resource policy used by direct HTML-to-PDF rendering.</summary>
public sealed class HtmlPdfResourcePolicySummary {
    private HtmlPdfResourcePolicySummary() {
    }

    /// <summary>True when an application-supplied asynchronous resource resolver is configured.</summary>
    public bool HasResourceResolver { get; private set; }

    /// <summary>True when installed host fonts may be discovered, read, and embedded.</summary>
    public bool AllowSystemFontEmbedding { get; private set; }

    /// <summary>True when local files may be resolved by the application-supplied resolver.</summary>
    public bool AllowLocalFileAccess { get; private set; }

    /// <summary>True when remote resources may be resolved by the application-supplied resolver.</summary>
    public bool AllowRemoteResourceResolution { get; private set; }

    /// <summary>True when inline data URI resources are accepted.</summary>
    public bool AllowDataUris { get; private set; }

    /// <summary>True when bounded resources already contained in a source package may be resolved.</summary>
    public bool AllowEmbeddedPackageResources { get; private set; }

    /// <summary>Maximum duration allowed for one resource request.</summary>
    public TimeSpan ResourceTimeout { get; private set; }

    /// <summary>Maximum asynchronous resource resolver concurrency.</summary>
    public int MaxConcurrentResourceLoads { get; private set; }

    /// <summary>Maximum bytes accepted from one resource.</summary>
    public long MaxResourceBytes { get; private set; }

    /// <summary>Maximum total bytes accepted by one conversion.</summary>
    public long MaxTotalResourceBytes { get; private set; }

    /// <summary>Maximum external resource count accepted by one conversion.</summary>
    public int MaxResourceCount { get; private set; }

    /// <summary>Maximum resolver invocations, including misses and failures.</summary>
    public int MaxResourceRequests { get; private set; }

    /// <summary>Maximum recursive stylesheet import depth.</summary>
    public int MaxStylesheetImportDepth { get; private set; }

    /// <summary>Allowed URL schemes when the URL policy restricts schemes.</summary>
    public IReadOnlyList<string> AllowedUrlSchemes { get; private set; } = Array.Empty<string>();

    /// <summary>Builds a detached resource-policy summary for the supplied options.</summary>
    public static HtmlPdfResourcePolicySummary From(HtmlPdfSaveOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        HtmlUrlPolicy urlPolicy = options.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();
        return new HtmlPdfResourcePolicySummary {
            HasResourceResolver = options.ResourceResolver != null,
            AllowSystemFontEmbedding = options.ResourcePolicy.AllowSystemFontEmbedding,
            AllowLocalFileAccess = options.ResourcePolicy.AllowLocalFileAccess,
            AllowRemoteResourceResolution = options.ResourcePolicy.AllowRemoteResourceResolution,
            AllowDataUris = options.ResourcePolicy.AllowDataUris,
            AllowEmbeddedPackageResources = options.ResourcePolicy.AllowEmbeddedPackageResources,
            ResourceTimeout = options.ResourceTimeout,
            MaxConcurrentResourceLoads = options.MaxConcurrentResourceLoads,
            MaxResourceBytes = options.MaxResourceBytes,
            MaxTotalResourceBytes = options.MaxTotalResourceBytes,
            MaxResourceCount = options.MaxResourceCount,
            MaxResourceRequests = options.MaxResourceRequests,
            MaxStylesheetImportDepth = options.MaxStylesheetImportDepth,
            AllowedUrlSchemes = urlPolicy.RestrictUrlSchemes
                ? CopySorted(urlPolicy.AllowedUrlSchemes)
                : Array.Empty<string>()
        };
    }

    private static IReadOnlyList<string> CopySorted(IEnumerable<string> values) {
        var copy = new List<string>();
        foreach (string value in values) {
            if (!string.IsNullOrWhiteSpace(value)) copy.Add(value);
        }
        copy.Sort(StringComparer.OrdinalIgnoreCase);
        return new ReadOnlyCollection<string>(copy);
    }
}
