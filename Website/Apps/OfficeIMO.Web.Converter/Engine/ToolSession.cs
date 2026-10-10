using System.Globalization;
using System.Text.Json;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>Worker-local state: staged inputs, the latest artifacts, and state a follow-up action needs.</summary>
internal sealed class ToolSession {
    internal const int MaxInputs = BrowserPdfToolService.MaxPdfFiles;
    private readonly SelectedDocument?[] _inputs = new SelectedDocument?[MaxInputs];
    private readonly List<BrowserConversionArtifact> _artifacts = [];
    // Keep PDF-specific types out of generic cache operations: text-only workers do not load the PDF assembly.
    private readonly Dictionary<string, object> _previews = new(StringComparer.Ordinal);

    internal BrowserConversionService Conversions { get; } = new();
    internal BrowserPdfToolService PdfTools { get; } = new();
    internal ConversionResult? LastConversion { get; set; }
    internal object? LastReview { get; set; }

    internal IReadOnlyList<SelectedDocument> Inputs => _inputs.Where(static input => input is not null).Select(static input => input!).ToArray();

    internal void Stage(int slot, byte[] bytes, string fileName) {
        if (slot < 0 || slot >= MaxInputs) throw new ArgumentOutOfRangeException(nameof(slot), $"Up to {MaxInputs} files can be staged.");
        if (bytes.LongLength > BrowserConversionService.MaxPackageBytes) {
            throw new InvalidDataException($"{fileName} is larger than the {ToolFormat.Bytes(BrowserConversionService.MaxPackageBytes)} browser limit.");
        }
        string extension = Path.GetExtension(fileName).ToLowerInvariant();
        _inputs[slot] = new SelectedDocument(fileName, extension, extension.TrimStart('.').ToUpperInvariant(), bytes.LongLength, bytes);
        DropPreview("input:" + slot.ToString(CultureInfo.InvariantCulture));
    }

    internal void ClearInputs() {
        Array.Clear(_inputs);
        _previews.Clear();
        _artifacts.Clear();
        LastConversion = null;
        LastReview = null;
    }

    internal SelectedDocument Input(int slot = 0) =>
        slot >= 0 && slot < MaxInputs && _inputs[slot] is { } input
            ? input
            : throw new InvalidOperationException("Add a file first.");

    internal ToolArtifact[] ReplaceArtifacts(params (BrowserConversionArtifact? Artifact, string Role)[] artifacts) {
        _artifacts.Clear();
        foreach (string key in _previews.Keys.Where(static key => key.StartsWith("artifact:", StringComparison.Ordinal)).ToArray()) _previews.Remove(key);
        return AppendArtifacts(artifacts);
    }

    internal ToolArtifact[] AppendArtifacts(params (BrowserConversionArtifact? Artifact, string Role)[] artifacts) {
        var added = new List<ToolArtifact>();
        foreach ((BrowserConversionArtifact? artifact, string role) in artifacts) {
            if (artifact is null) continue;
            _artifacts.Add(artifact);
            added.Add(new ToolArtifact(_artifacts.Count - 1, artifact.FileName, artifact.ContentType, artifact.Bytes.LongLength, role));
        }
        return added.ToArray();
    }

    internal byte[] Artifact(int index) =>
        index >= 0 && index < _artifacts.Count ? _artifacts[index].Bytes : throw new ArgumentOutOfRangeException(nameof(index));

    internal BrowserPdfPreview Preview(string source, int index) {
        string key = source + ":" + index.ToString(CultureInfo.InvariantCulture);
        if (_previews.TryGetValue(key, out object? cached)) return (BrowserPdfPreview)cached;
        byte[] bytes = source switch {
            "input" => Input(index).Bytes,
            "artifact" => Artifact(index),
            _ => throw new ArgumentOutOfRangeException(nameof(source))
        };
        var preview = new BrowserPdfPreview(bytes);
        _previews[key] = preview;
        return preview;
    }

    private void DropPreview(string key) => _previews.Remove(key);
}

/// <summary>Flat option bag sent by the tool shell.</summary>
internal sealed class ToolOptions(Dictionary<string, string> values) {
    internal static ToolOptions Parse(string? json) {
        var parsed = string.IsNullOrWhiteSpace(json)
            ? new Dictionary<string, string>(StringComparer.Ordinal)
            : JsonSerializer.Deserialize(json, EngineJsonContext.Default.DictionaryStringString) ?? [];
        foreach (string key in parsed.Keys.Where(static key => key.StartsWith("_utf16:", StringComparison.Ordinal)).ToArray()) {
            string target = key[7..];
            if (target.Length == 0 || parsed.ContainsKey(target)) throw new ArgumentException("The UTF-16 option has a duplicate or missing name.");
            parsed[target] = Utf16Transport.Decode(parsed[key]);
            parsed.Remove(key);
        }
        return new ToolOptions(parsed);
    }

    internal string Text(string key, string fallback = "") =>
        values.TryGetValue(key, out string? value) && value is not null ? value : fallback;

    internal bool Flag(string key) =>
        values.TryGetValue(key, out string? value) && (value == "true" || value == "on" || value == "1");

    internal int Number(string key, int fallback) {
        if (!values.TryGetValue(key, out string? value)) return fallback;
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int number)) {
            throw new ArgumentException($"The '{key}' option must be a whole number.", key);
        }
        return number;
    }

    internal IReadOnlySet<string> List(string key) =>
        Text(key).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToHashSet(StringComparer.Ordinal);
}

internal static class ToolFormat {
    internal static string Bytes(long bytes) {
        string[] units = ["B", "KB", "MB", "GB"];
        double value = bytes;
        int unit = 0;
        while (value >= 1024 && unit < units.Length - 1) { value /= 1024; unit++; }
        return unit == 0 ? $"{bytes} B" : value.ToString(value >= 100 ? "0" : "0.#", CultureInfo.InvariantCulture) + " " + units[unit];
    }

    internal static string Duration(long milliseconds) =>
        milliseconds < 1000 ? $"{milliseconds} ms" : (milliseconds / 1000d).ToString("0.0", CultureInfo.InvariantCulture) + " s";

    internal static string Count(int count, string singular, string? plural = null) =>
        count.ToString(CultureInfo.InvariantCulture) + " " + (count == 1 ? singular : plural ?? singular + "s");

    internal static string Pages(IEnumerable<int> pages) {
        int[] list = pages.Distinct().Order().ToArray();
        if (list.Length == 0) return string.Empty;
        string joined = list.Length switch {
            1 => list[0].ToString(CultureInfo.InvariantCulture),
            2 => $"{list[0]} and {list[1]}",
            <= 6 => string.Join(", ", list[..^1]) + " and " + list[^1],
            _ => string.Join(", ", list[..5]) + $" and {list.Length - 5} more"
        };
        return (list.Length == 1 ? "page " : "pages ") + joined;
    }

    /// <summary>Turns engine identifiers such as RightToLeftOverride into readable words.</summary>
    internal static string Words(string identifier) {
        if (string.IsNullOrEmpty(identifier)) return identifier;
        var builder = new System.Text.StringBuilder(identifier.Length + 8);
        for (int index = 0; index < identifier.Length; index++) {
            char current = identifier[index];
            if (index > 0 && char.IsUpper(current) && (char.IsLower(identifier[index - 1]) || (index + 1 < identifier.Length && char.IsLower(identifier[index + 1])))) {
                builder.Append(' ');
                builder.Append(char.ToLowerInvariant(current));
            } else {
                builder.Append(index == 0 ? char.ToUpperInvariant(current) : current);
            }
        }
        return builder.ToString();
    }
}
