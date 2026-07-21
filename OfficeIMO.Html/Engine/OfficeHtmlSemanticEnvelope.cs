using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Owns the versioned OfficeIMO semantic HTML envelope contract.</summary>
public static class OfficeHtmlSemanticEnvelope {
    /// <summary>Current semantic-envelope schema version emitted by OfficeIMO adapters.</summary>
    public const string CurrentSchemaVersion = "1";

    /// <summary>HTML attribute carrying the semantic-envelope schema version.</summary>
    public const string SchemaVersionAttribute = "data-officeimo-schema-version";

    /// <summary>Appends canonical source, profile, and schema attributes to an open start tag.</summary>
    public static void AppendRootAttributes(StringBuilder builder, string source, string profile) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("Semantic source cannot be empty.", nameof(source));
        if (string.IsNullOrWhiteSpace(profile)) throw new ArgumentException("Semantic profile cannot be empty.", nameof(profile));
        builder.Append(" data-officeimo-source=\"")
            .Append(OfficeHtmlText.EscapeAttribute(source.Trim().ToLowerInvariant()))
            .Append("\" data-officeimo-profile=\"")
            .Append(OfficeHtmlText.EscapeAttribute(profile.Trim()))
            .Append("\" ")
            .Append(SchemaVersionAttribute)
            .Append("=\"")
            .Append(CurrentSchemaVersion)
            .Append('"');
    }

    /// <summary>Inspects a parsed document using the shared compatibility rules.</summary>
    public static OfficeHtmlSemanticEnvelopeInfo Inspect(IHtmlDocument document, string expectedSource) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(expectedSource)) throw new ArgumentException("Expected source cannot be empty.", nameof(expectedSource));

        string normalizedSource = expectedSource.Trim().ToLowerInvariant();
        IElement? root = document.QuerySelector("main.officeimo-document[data-officeimo-source]");
        if (root == null) return OfficeHtmlSemanticEnvelopeInfo.Missing(normalizedSource);

        string actualSource = (root.GetAttribute("data-officeimo-source") ?? string.Empty).Trim().ToLowerInvariant();
        string version = (root.GetAttribute(SchemaVersionAttribute) ?? string.Empty).Trim();
        bool sourceMatches = string.Equals(actualSource, normalizedSource, StringComparison.OrdinalIgnoreCase);
        if (version.Length == 0) return OfficeHtmlSemanticEnvelopeInfo.Legacy(normalizedSource, actualSource, sourceMatches, root);
        return new OfficeHtmlSemanticEnvelopeInfo(
            isPresent: true,
            isLegacy: false,
            isSupported: sourceMatches && string.Equals(version, CurrentSchemaVersion, StringComparison.Ordinal),
            sourceMatches,
            normalizedSource,
            actualSource,
            version,
            root);
    }

    /// <summary>Selects semantic containers owned by the inspected envelope, excluding nested envelopes.</summary>
    internal static IReadOnlyList<IElement> SelectOwnedContainers(
        IHtmlDocument document,
        OfficeHtmlSemanticEnvelopeInfo envelope,
        string selector) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (envelope == null) throw new ArgumentNullException(nameof(envelope));
        if (string.IsNullOrWhiteSpace(selector)) throw new ArgumentException("A container selector is required.", nameof(selector));

        IElement? root = envelope.RootElement ?? document.DocumentElement;
        if (root == null) return Array.Empty<IElement>();
        if (envelope.RootElement == null) return root.QuerySelectorAll(selector).ToList();
        return root.QuerySelectorAll(selector)
            .Where(element => ReferenceEquals(FindEnvelopeRoot(element), root))
            .ToList();
    }

    private static IElement? FindEnvelopeRoot(IElement element) {
        for (IElement? current = element.ParentElement; current != null; current = current.ParentElement) {
            if (current.Matches("main.officeimo-document[data-officeimo-source]")) return current;
        }
        return null;
    }
}

/// <summary>Describes semantic-envelope presence and compatibility.</summary>
public sealed class OfficeHtmlSemanticEnvelopeInfo {
    internal OfficeHtmlSemanticEnvelopeInfo(
        bool isPresent,
        bool isLegacy,
        bool isSupported,
        bool sourceMatches,
        string expectedSource,
        string actualSource,
        string schemaVersion,
        IElement? rootElement) {
        IsPresent = isPresent;
        IsLegacy = isLegacy;
        IsSupported = isSupported;
        SourceMatches = sourceMatches;
        ExpectedSource = expectedSource;
        ActualSource = actualSource;
        SchemaVersion = schemaVersion;
        RootElement = rootElement;
    }

    internal static OfficeHtmlSemanticEnvelopeInfo Missing(string expectedSource) =>
        new OfficeHtmlSemanticEnvelopeInfo(false, false, true, true, expectedSource, string.Empty, string.Empty, null);

    internal static OfficeHtmlSemanticEnvelopeInfo Legacy(string expectedSource, string actualSource, bool sourceMatches, IElement rootElement) =>
        new OfficeHtmlSemanticEnvelopeInfo(true, true, sourceMatches, sourceMatches, expectedSource, actualSource, string.Empty, rootElement);

    /// <summary>The exact envelope root that was inspected, or <see langword="null"/> when no envelope was found.</summary>
    internal IElement? RootElement { get; }

    /// <summary>Whether a canonical OfficeIMO root envelope was present.</summary>
    public bool IsPresent { get; }

    /// <summary>Whether the input predates explicit schema versioning.</summary>
    public bool IsLegacy { get; }

    /// <summary>Whether the source and schema version are supported by the requested adapter.</summary>
    public bool IsSupported { get; }

    /// <summary>Whether the envelope source matches the requested adapter.</summary>
    public bool SourceMatches { get; }

    /// <summary>Source identifier expected by the requested adapter.</summary>
    public string ExpectedSource { get; }

    /// <summary>Source identifier declared by the envelope.</summary>
    public string ActualSource { get; }

    /// <summary>Declared schema version, or an empty string for legacy input.</summary>
    public string SchemaVersion { get; }
}
