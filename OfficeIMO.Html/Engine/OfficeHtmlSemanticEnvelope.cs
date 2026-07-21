using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Owns the versioned OfficeIMO semantic HTML envelope contract.</summary>
public static class OfficeHtmlSemanticEnvelope {
    /// <summary>Current semantic-envelope schema version emitted by OfficeIMO adapters.</summary>
    public const string CurrentSchemaVersion = "2";

    /// <summary>Previous explicit schema version accepted for backward-compatible imports.</summary>
    public const string PreviousSchemaVersion = "1";

    /// <summary>HTML attribute carrying the semantic-envelope schema version.</summary>
    public const string SchemaVersionAttribute = "data-officeimo-schema-version";

    /// <summary>HTML attribute declaring whether restoration metadata is public-safe or caller-trusted.</summary>
    public const string RestorationAttribute = "data-officeimo-restoration";

    /// <summary>HTML attribute declaring the validation state of public semantic metadata.</summary>
    public const string PublicSemanticsAttribute = "data-officeimo-public-semantics";

    /// <summary>Validated public semantic metadata marker emitted by schema version 2.</summary>
    public const string PublicSemanticsSafe = "safe";

    /// <summary>Restoration mode that may be used at an untrusted input boundary.</summary>
    public const string PublicSafeRestoration = "public-safe";

    /// <summary>Restoration mode that requires an explicitly trusted input boundary.</summary>
    public const string TrustedTargetRestoration = "trusted-target";

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
            .Append("\" ")
            .Append(PublicSemanticsAttribute)
            .Append("=\"")
            .Append(PublicSemanticsSafe)
            .Append("\" ")
            .Append(RestorationAttribute)
            .Append("=\"")
            .Append(PublicSafeRestoration)
            .Append("\"");
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
        string restoration = (root.GetAttribute(RestorationAttribute) ?? string.Empty).Trim().ToLowerInvariant();
        string publicSemantics = (root.GetAttribute(PublicSemanticsAttribute) ?? string.Empty).Trim().ToLowerInvariant();
        bool sourceMatches = string.Equals(actualSource, normalizedSource, StringComparison.OrdinalIgnoreCase);
        if (version.Length == 0) return OfficeHtmlSemanticEnvelopeInfo.Legacy(normalizedSource, actualSource, sourceMatches, root);
        bool contractSupported = IsSupportedVersion(version)
            && (string.Equals(version, PreviousSchemaVersion, StringComparison.Ordinal)
                || IsSupportedV2Contract(restoration, publicSemantics));
        return new OfficeHtmlSemanticEnvelopeInfo(
            isPresent: true,
            isLegacy: false,
            isSupported: sourceMatches && contractSupported,
            sourceMatches,
            normalizedSource,
            actualSource,
            version,
            restoration,
            publicSemantics,
            root);
    }

    /// <summary>Returns whether an explicit envelope schema version is accepted.</summary>
    public static bool IsSupportedVersion(string? version) =>
        string.Equals(version, CurrentSchemaVersion, StringComparison.Ordinal)
        || string.Equals(version, PreviousSchemaVersion, StringComparison.Ordinal);

    private static bool IsSupportedV2Contract(string restoration, string publicSemantics) =>
        string.Equals(publicSemantics, PublicSemanticsSafe, StringComparison.OrdinalIgnoreCase)
        && (string.Equals(restoration, PublicSafeRestoration, StringComparison.OrdinalIgnoreCase)
            || string.Equals(restoration, TrustedTargetRestoration, StringComparison.OrdinalIgnoreCase));

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
        string restorationMode,
        string publicSemantics,
        IElement? rootElement) {
        IsPresent = isPresent;
        IsLegacy = isLegacy;
        IsSupported = isSupported;
        SourceMatches = sourceMatches;
        ExpectedSource = expectedSource;
        ActualSource = actualSource;
        SchemaVersion = schemaVersion;
        RestorationMode = restorationMode;
        PublicSemantics = publicSemantics;
        RootElement = rootElement;
    }

    internal static OfficeHtmlSemanticEnvelopeInfo Missing(string expectedSource) =>
        new OfficeHtmlSemanticEnvelopeInfo(false, false, true, true, expectedSource, string.Empty, string.Empty, string.Empty, string.Empty, null);

    internal static OfficeHtmlSemanticEnvelopeInfo Legacy(string expectedSource, string actualSource, bool sourceMatches, IElement rootElement) =>
        new OfficeHtmlSemanticEnvelopeInfo(true, true, sourceMatches, sourceMatches, expectedSource, actualSource, string.Empty, string.Empty, string.Empty, rootElement);

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

    /// <summary>Declared restoration mode. Version 1 and legacy envelopes return an empty value.</summary>
    public string RestorationMode { get; }

    /// <summary>Declared public-semantics validation marker. Version 1 and legacy envelopes return an empty value.</summary>
    public string PublicSemantics { get; }

    /// <summary>Whether target-specific restoration requires caller-trusted input.</summary>
    public bool RequiresTrustedRestoration =>
        IsSupported
        && string.Equals(SchemaVersion, OfficeHtmlSemanticEnvelope.CurrentSchemaVersion, StringComparison.Ordinal)
        && string.Equals(RestorationMode, OfficeHtmlSemanticEnvelope.TrustedTargetRestoration, StringComparison.OrdinalIgnoreCase);

    /// <summary>Whether target-specific restoration is allowed at the caller-assigned trust boundary.</summary>
    public bool CanRestoreTargetSpecific(HtmlInputTrust trust) =>
        IsSupported
        && (!RequiresTrustedRestoration || trust == HtmlInputTrust.Trusted);
}
