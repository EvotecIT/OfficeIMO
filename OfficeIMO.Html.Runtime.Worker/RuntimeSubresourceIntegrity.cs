using System.Security.Cryptography;
using AngleSharp.Io;

namespace OfficeIMO.Html.Runtime.Worker;

// One owned verifier serves retained DOM loads, modulepreload and import-map
// integrity metadata. It can be replaced with the rest of the provider boundary.
internal sealed class RuntimeSubresourceIntegrity(int maximumMetadataCharacters) : IIntegrityProvider {
    public bool IsSatisfied(byte[] content, string integrity) {
        if (integrity.Length > maximumMetadataCharacters) return false;
        var candidates = Parse(integrity);
        if (candidates.Count == 0) return true;
        int strongest = candidates.Max(candidate => candidate.Strength);
        foreach (var candidate in candidates.Where(candidate => candidate.Strength == strongest)) {
            byte[] actual = candidate.Strength switch {
                1 => SHA256.HashData(content),
                2 => SHA384.HashData(content),
                _ => SHA512.HashData(content)
            };
            if (candidate.Digest != null && CryptographicOperations.FixedTimeEquals(actual, candidate.Digest)) return true;
        }
        return false;
    }

    private static List<Candidate> Parse(string metadata) {
        var result = new List<Candidate>();
        foreach (string token in metadata.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
            int separator = token.IndexOf('-');
            if (separator <= 0) continue;
            int strength = token[..separator].ToLowerInvariant() switch { "sha256" => 1, "sha384" => 2, "sha512" => 3, _ => 0 };
            if (strength == 0) continue;
            string encoded = token[(separator + 1)..].Split('?', 2)[0];
            try { result.Add(new(strength, Convert.FromBase64String(encoded))); }
            catch (FormatException) { result.Add(new(strength, null)); }
        }
        return result;
    }

    private sealed record Candidate(int Strength, byte[]? Digest);
}
