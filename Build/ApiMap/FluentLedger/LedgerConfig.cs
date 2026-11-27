using System.Text.Json;

namespace OfficeIMO.Build.ApiMap.Fluent;

internal sealed record FamilyOverride(string Old, string? Package, string New, string? Treatment, string Reason);
internal sealed record PropertyRename(string Receiver, string? Package, string Old, string New, string Reason);
internal sealed record Retirement(string Receiver, string Member, string Replacement, string Reason);
internal sealed record DecisionRule(string Receiver, string[] Members, string Code, string Proposal, string Reason);

/// <summary>The reviewable classification inputs: package scope, #3344 family overrides and verb lists.</summary>
internal sealed record LedgerConfig(
    string Issue,
    string[] Frameworks,
    string[] Packages,
    FamilyOverride[] Families,
    PropertyRename[] PropertyRenames,
    Retirement[] Retire,
    DecisionRule[] Decisions,
    string[] SetPrefixes,
    string[] CopyPrefixes,
    string[] AddPrefixes,
    string[] LifecycleNames,
    string[] ConversionPrefixes,
    string[] SelectorPrefixes,
    string[] ActionPrefixes,
    string[] ObjectContract) {
    internal static LedgerConfig Load(string path) =>
        JsonSerializer.Deserialize<LedgerConfig>(File.ReadAllText(path), new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
        ?? throw new InvalidDataException("The ledger configuration is empty: " + path);

    /// <summary>Matches a receiver pattern: an exact type name, <c>*</c>, or a namespace prefix ending in <c>.*</c>.</summary>
    internal static bool ReceiverMatches(string pattern, string receiver) =>
        pattern == "*" || pattern == receiver
        || pattern.EndsWith(".*", StringComparison.Ordinal) && receiver.StartsWith(pattern[..^1], StringComparison.Ordinal);

    /// <summary>A verb prefix only matches at a word boundary, so <c>Is</c> does not match <c>Insert</c>.</summary>
    internal static string? MatchPrefix(IEnumerable<string> prefixes, string name) {
        foreach (string prefix in prefixes.OrderByDescending(prefix => prefix.Length)) {
            if (!name.StartsWith(prefix, StringComparison.Ordinal)) continue;
            if (name.Length == prefix.Length || char.IsUpper(name[prefix.Length]) || char.IsDigit(name[prefix.Length])) return prefix;
        }
        return null;
    }
}
