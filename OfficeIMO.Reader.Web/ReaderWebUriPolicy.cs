namespace OfficeIMO.Reader.Web;

internal static class ReaderWebUriPolicy {
    internal static void Validate(Uri uri, ReaderWebOptions options) {
        if (uri == null) throw new ArgumentNullException(nameof(uri));
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
        if (!options.AllowLocalhostAndNonPublicIpLiterals && IsLocalhostOrNonPublicIpLiteral(uri, host)) {
            throw new ReaderWebPolicyException(uri, "Reader Web rejected a localhost name or a loopback, private, link-local, or non-routable IP literal.");
        }
    }

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

    private static bool IsLocalhostOrNonPublicIpLiteral(Uri uri, string host) {
        if (uri.IsLoopback ||
            string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase) ||
            host.EndsWith(".localhost", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }
        if (!IPAddress.TryParse(host, out IPAddress? address)) return false;
        if (address.IsIPv4MappedToIPv6) address = address.MapToIPv4();
        if (IPAddress.IsLoopback(address) ||
            address.Equals(IPAddress.Any) ||
            address.Equals(IPAddress.IPv6Any) ||
            address.Equals(IPAddress.None) ||
            address.Equals(IPAddress.IPv6None)) {
            return true;
        }

        byte[] bytes = address.GetAddressBytes();
        if (address.AddressFamily == AddressFamily.InterNetwork) {
            return IsPrivateOrNonRoutableIpv4(bytes);
        }
        if (address.AddressFamily == AddressFamily.InterNetworkV6) {
            return address.IsIPv6LinkLocal ||
                address.IsIPv6SiteLocal ||
                address.IsIPv6Multicast ||
                (bytes[0] & 0xFE) == 0xFC ||
                IsDiscardOnlyIpv6(bytes) ||
                IsLocalUseNat64(bytes) ||
                IsWellKnownNat64WithPrivateIpv4(bytes) ||
                IsIpv4TranslatedWithPrivateIpv4(bytes) ||
                IsSixToFourWithPrivateIpv4(bytes) ||
                (bytes[0] == 0x20 && bytes[1] == 0x01 && bytes[2] == 0x0D && bytes[3] == 0xB8);
        }
        return true;
    }

    private static bool IsPrivateOrNonRoutableIpv4(byte[] bytes) {
        byte first = bytes[0];
        byte second = bytes[1];
        byte third = bytes[2];
        if (first == 0 || first == 10 || first == 127 || first >= 224) return true;
        if (first == 100 && second >= 64 && second <= 127) return true;
        if (first == 169 && second == 254) return true;
        if (first == 172 && second >= 16 && second <= 31) return true;
        if (first == 192 && second == 168) return true;
        if (first == 192 && second == 0 && third == 0) {
            byte fourth = bytes[3];
            return fourth != 9 && fourth != 10;
        }
        if (first == 192 && second == 0 && third == 2) return true;
        if (first == 192 && second == 88 && third == 99) return true;
        if (first == 198 && (second == 18 || second == 19 || (second == 51 && third == 100))) return true;
        if (first == 203 && second == 0 && third == 113) return true;
        return false;
    }

    private static bool IsDiscardOnlyIpv6(byte[] bytes) {
        if (bytes.Length != 16 || bytes[0] != 0x01) return false;
        for (int index = 1; index < 8; index++) {
            if (bytes[index] != 0x00) return false;
        }
        return true;
    }

    private static bool IsLocalUseNat64(byte[] bytes) {
        return bytes.Length == 16 &&
            bytes[0] == 0x00 && bytes[1] == 0x64 &&
            bytes[2] == 0xFF && bytes[3] == 0x9B &&
            bytes[4] == 0x00 && bytes[5] == 0x01;
    }

    private static bool IsWellKnownNat64WithPrivateIpv4(byte[] bytes) {
        if (bytes.Length != 16 ||
            bytes[0] != 0x00 || bytes[1] != 0x64 ||
            bytes[2] != 0xFF || bytes[3] != 0x9B) {
            return false;
        }
        for (int index = 4; index < 12; index++) {
            if (bytes[index] != 0x00) return false;
        }
        return IsPrivateOrNonRoutableIpv4(new[] { bytes[12], bytes[13], bytes[14], bytes[15] });
    }

    private static bool IsIpv4TranslatedWithPrivateIpv4(byte[] bytes) {
        if (bytes.Length != 16 ||
            bytes[0] != 0x00 || bytes[1] != 0x00 ||
            bytes[2] != 0x00 || bytes[3] != 0x00 ||
            bytes[4] != 0x00 || bytes[5] != 0x00 ||
            bytes[6] != 0x00 || bytes[7] != 0x00 ||
            bytes[8] != 0xFF || bytes[9] != 0xFF ||
            bytes[10] != 0x00 || bytes[11] != 0x00) {
            return false;
        }
        return IsPrivateOrNonRoutableIpv4(new[] { bytes[12], bytes[13], bytes[14], bytes[15] });
    }

    private static bool IsSixToFourWithPrivateIpv4(byte[] bytes) {
        return bytes.Length == 16 &&
            bytes[0] == 0x20 && bytes[1] == 0x02 &&
            IsPrivateOrNonRoutableIpv4(new[] { bytes[2], bytes[3], bytes[4], bytes[5] });
    }
}
