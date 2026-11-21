using OfficeIMO.Core.Internal;

namespace OfficeIMO.Reader.Web;

internal static class ReaderWebUriPolicy {
    internal static void Validate(Uri uri, ReaderWebOptions options, bool screenHostNames = true) {
        ArgumentNullException.ThrowIfNull(uri);
        if (!uri.IsAbsoluteUri) {
            throw new ReaderWebPolicyException(uri, "Reader Web requires an absolute URI.");
        }
        if (!string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) {
            throw new ReaderWebPolicyException(uri, "Reader Web permits only HTTP and HTTPS URIs.");
        }
        if (!string.IsNullOrEmpty(uri.UserInfo)) {
            throw new ReaderWebPolicyException(uri, "Reader Web rejects credentials embedded in a URI.");
        }

        string host = uri.IdnHost.TrimEnd('.').ToLowerInvariant();
        if (options.AllowedHosts.Count > 0 && !IsAllowedHost(host, options)) {
            throw new ReaderWebPolicyException(uri, "Reader Web rejected a host outside the configured allowlist.");
        }
        if (!options.AllowPrivateNetworkAddresses && screenHostNames) {
            try {
                OfficeNetworkAddressPolicy.ThrowIfNonPublicHostName(uri);
            } catch (OfficeNetworkAddressException exception) {
                throw Rejected(uri, exception);
            }
        }
    }

    internal static ReaderWebPolicyException Rejected(Uri uri, OfficeNetworkAddressException exception) =>
        new(uri, "Reader Web rejected a localhost name or a loopback, private, link-local, or other non-public network address. " + exception.Message);

    private static bool IsAllowedHost(string host, ReaderWebOptions options) {
        for (int index = 0; index < options.AllowedHosts.Count; index++) {
            string allowed = options.AllowedHosts[index];
            if (string.Equals(host, allowed, StringComparison.OrdinalIgnoreCase)) return true;
            if (options.AllowSubdomains &&
                host.Length > allowed.Length &&
                host.EndsWith("." + allowed, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }
        return false;
    }
}
