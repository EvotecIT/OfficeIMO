using System.Text.Json.Serialization;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>Uniform result returned to the tool shell for every browser tool and action.</summary>
internal sealed record ToolResultDocument(
    bool Ok,
    ToolVerdict Verdict,
    ToolFact[] Facts,
    ToolItem[] Items,
    ToolArtifact[] Artifacts,
    ToolPreview? Preview,
    long ElapsedMilliseconds,
    long? PeakRetainedBytes = null,
    string[]? Needs = null) {
    internal static ToolResultDocument Failure(string title, string detail, long elapsedMilliseconds = 0) =>
        new(false, new ToolVerdict(ToolTone.Bad, title, detail), [], [], [], null, elapsedMilliseconds);

    /// <summary>Asks the worker to load these lazy assemblies (for example "OfficeIMO.Web.Fonts.Fallback.wasm") and run the same call again.</summary>
    internal static ToolResultDocument NeedsAssemblies(params string[] assemblies) =>
        Failure("Loading more of the engine", "This document needs extra fonts. They are downloading now.") with { Needs = assemblies };
}

/// <summary>Thrown before any work starts when the run needs a lazy assembly the worker hasn't loaded yet.</summary>
internal sealed class EngineNeedsAssemblyException(params string[] assemblies)
    : Exception("The engine needs " + string.Join(", ", assemblies) + ".") {
    internal string[] Assemblies { get; } = assemblies;
}

/// <summary>One-sentence answer shown first. Tone is good, warn, bad, or info.</summary>
internal sealed record ToolVerdict(string Tone, string Title, string Detail);

internal sealed record ToolFact(string Label, string Value, string? Tone = null);

/// <summary>A named finding, warning, or change. State drives the badge: found, removed, kept, warning, info, good, bad.</summary>
internal sealed record ToolItem(
    string Id,
    string Title,
    string Detail,
    string State,
    string? Group = null,
    int? Page = null,
    bool Selectable = false,
    bool Selected = false,
    int? Start = null,
    int? Length = null);

/// <summary>A file the shell can download. Role is primary, report, overlay, support, or preview.</summary>
internal sealed record ToolArtifact(int Index, string FileName, string ContentType, long Bytes, string Role);

/// <summary>How the shell previews the primary output: pdf, image, html, text, or none.</summary>
internal sealed record ToolPreview(string Kind, int? ArtifactIndex = null, int? PageCount = null, string? Html = null, string? Text = null) {
    // JSON strings cannot represent lone surrogates. The worker restores this exact code-unit payload.
    public string? TextUtf16 => Utf16Transport.EncodeIfUnpaired(Text);
}

internal sealed record PdfProbeDocument(bool Ok, int PageCount, bool Encrypted, bool NeedsPassword, bool CanManipulatePages, string? Error);

internal static class ToolTone {
    internal const string Good = "good";
    internal const string Warn = "warn";
    internal const string Bad = "bad";
    internal const string Info = "info";
}

internal static class ToolState {
    internal const string Found = "found";
    internal const string Removed = "removed";
    internal const string Kept = "kept";
    internal const string Warning = "warning";
    internal const string Info = "info";
    internal const string Good = "good";
    internal const string Bad = "bad";
    /// <summary>A plain explanation row; the page shows no status badge for it.</summary>
    internal const string Detail = "detail";
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(ToolResultDocument))]
[JsonSerializable(typeof(PdfProbeDocument))]
[JsonSerializable(typeof(Dictionary<string, string>))]
internal sealed partial class EngineJsonContext : JsonSerializerContext;
