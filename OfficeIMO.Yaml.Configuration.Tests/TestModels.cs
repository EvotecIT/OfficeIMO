using Microsoft.Extensions.Configuration;
using Xunit;

namespace OfficeIMO.Yaml.Configuration.Tests;

public sealed class AppOptions {
    public string Name { get; set; } = string.Empty;
    public int Port { get; set; }
    public bool Enabled { get; set; }
    public DateTimeOffset Started { get; set; }
    public DatabaseOptions Database { get; set; } = new DatabaseOptions();
    public List<string> Hosts { get; set; } = new List<string>();
    public List<EndpointOptions> Endpoints { get; set; } = new List<EndpointOptions>();
    public Dictionary<string, int> Limits { get; set; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
}

public sealed class DatabaseOptions {
    public string Host { get; set; } = string.Empty;
    public TimeSpan Timeout { get; set; }
}

public sealed class EndpointOptions {
    public string Name { get; set; } = string.Empty;
    public int Weight { get; set; }
}

internal static class ConfigurationParity {
    // The JSON provider spells booleans True and False; YAML and TOML keep the file's spelling. Both bind the same way, so
    // booleans compare by value and every other value compares exactly.
    public static void AssertEquivalent(IConfiguration expected, IConfiguration actual) {
        Dictionary<string, string?> left = expected.AsEnumerable().ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        Dictionary<string, string?> right = actual.AsEnumerable().ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);

        Assert.Equal(left.Keys.OrderBy(key => key, StringComparer.Ordinal), right.Keys.OrderBy(key => key, StringComparer.Ordinal));
        foreach (KeyValuePair<string, string?> pair in left) {
            string? value = right[pair.Key];
            if (bool.TryParse(pair.Value, out bool expectedFlag) && bool.TryParse(value, out bool actualFlag)) {
                Assert.Equal(expectedFlag, actualFlag);
            } else {
                Assert.True(pair.Value == value, $"'{pair.Key}': expected '{pair.Value}', found '{value}'.");
            }
        }
    }
}
