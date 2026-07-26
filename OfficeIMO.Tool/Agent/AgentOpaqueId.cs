using System.Text;

namespace OfficeIMO.Tool.Agent;

internal static class AgentOpaqueId {
    internal static string Encode(string kind, string value) {
        ArgumentException.ThrowIfNullOrWhiteSpace(kind);
        ArgumentNullException.ThrowIfNull(value);
        return kind + ":" + Convert.ToBase64String(Encoding.UTF8.GetBytes(value))
            .TrimEnd('=')
            .Replace('+', '-')
            .Replace('/', '_');
    }

    internal static (string Kind, string Value) Decode(string id) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new AgentUsageException("A non-empty result id is required.");
        }
        int separator = id.IndexOf(':');
        if (separator <= 0 || separator == id.Length - 1) {
            throw new AgentUsageException("Result id is not a valid OfficeIMO opaque id.");
        }
        string encoded = id.Substring(separator + 1)
            .Replace('-', '+')
            .Replace('_', '/');
        encoded = encoded.PadRight(encoded.Length + ((4 - encoded.Length % 4) % 4), '=');
        try {
            return (id.Substring(0, separator), Encoding.UTF8.GetString(Convert.FromBase64String(encoded)));
        } catch (FormatException exception) {
            throw new AgentUsageException("Result id is not a valid OfficeIMO opaque id.", exception);
        }
    }
}
