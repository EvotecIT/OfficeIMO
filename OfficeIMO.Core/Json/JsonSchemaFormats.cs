using System.Globalization;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Text.RegularExpressions;

namespace OfficeIMO.Json;

/// <summary>Checks for the <c>format</c> values the validator asserts. Unknown formats return null.</summary>
internal static class JsonSchemaFormats {
    public static bool? Check(string format, string value, TimeSpan regexTimeout) => format switch {
        "date-time" => IsDateTime(value),
        "date" => IsDate(value),
        "time" => IsTime(value),
        "duration" => IsDuration(value),
        "email" => IsEmail(value, international: false),
        "idn-email" => IsEmail(value, international: true),
        "hostname" => JsonSchemaHostnames.IsHostname(value),
        "idn-hostname" => JsonSchemaHostnames.IsIdnHostname(value),
        "ipv4" => IsIPv4(value),
        "ipv6" => IsIPv6(value),
        "uri" => IsUri(value, absolute: true, international: false),
        "uri-reference" => IsUri(value, absolute: false, international: false),
        "iri" => IsUri(value, absolute: true, international: true),
        "iri-reference" => IsUri(value, absolute: false, international: true),
        "uri-template" => IsUriTemplate(value),
        "uuid" => IsUuid(value),
        "regex" => IsRegex(value, regexTimeout),
        "json-pointer" => IsJsonPointer(value),
        "relative-json-pointer" => IsRelativeJsonPointer(value),
        _ => null
    };

    #region Dates and times (RFC 3339)

    private static bool IsDateTime(string value) {
        int t = value.IndexOfAny(new[] { 'T', 't' });
        return t == 10 && IsDate(value.Substring(0, t)) && IsTime(value.Substring(t + 1));
    }

    private static bool IsDate(string value) {
        if (value.Length != 10 || value[4] != '-' || value[7] != '-') return false;
        if (!TryDigits(value, 0, 4, out int year) || !TryDigits(value, 5, 2, out int month) || !TryDigits(value, 8, 2, out int day)) return false;
        return month >= 1 && month <= 12 && day >= 1 && day <= DateTime.DaysInMonth(year == 0 ? 2000 : year, month);
    }

    private static bool IsTime(string value) {
        // partial-time time-offset: HH:MM:SS[.frac](Z|+HH:MM|-HH:MM)
        if (value.Length < 9 || value[2] != ':' || value[5] != ':') return false;
        if (!TryDigits(value, 0, 2, out int hour) || !TryDigits(value, 3, 2, out int minute) || !TryDigits(value, 6, 2, out int second)) return false;
        int i = 8;
        if (i < value.Length && value[i] == '.') {
            int start = ++i;
            while (i < value.Length && IsDigit(value[i])) i++;
            if (i == start) return false;
        }
        if (i >= value.Length) return false;
        int offsetMinutes;
        if (value[i] is 'Z' or 'z') {
            if (i + 1 != value.Length) return false;
            offsetMinutes = 0;
        } else if (value[i] is '+' or '-') {
            if (value.Length - i != 6 || value[i + 3] != ':') return false;
            if (!TryDigits(value, i + 1, 2, out int offsetHour) || !TryDigits(value, i + 4, 2, out int offsetMinute) || offsetHour > 23 || offsetMinute > 59) return false;
            offsetMinutes = (offsetHour * 60 + offsetMinute) * (value[i] == '-' ? -1 : 1);
        } else {
            return false;
        }
        if (hour > 23 || minute > 59 || second > 60) return false;
        if (second == 60) {
            // A leap second is only valid at 23:59:60 UTC.
            int utc = ((hour * 60 + minute - offsetMinutes) % 1440 + 1440) % 1440;
            return utc == 23 * 60 + 59;
        }
        return true;
    }

    private static bool IsDuration(string value) {
        // RFC 3339 appendix A: P(n)W, or P with date units Y, M, D and time units H, M, S after T. Each part is a
        // run of consecutive units (Y-M-D, M-D, D; H-M-S, M-S, S), so P1Y2D and PT1H2S are invalid.
        if (value.Length < 2 || value[0] != 'P') return false;
        int i = 1;
        if (TryDurationElement(value, ref i, out char weekUnit) && weekUnit == 'W') return i == value.Length;
        i = 1;
        int t = value.IndexOf('T');
        string date = t < 0 ? value.Substring(1) : value.Substring(1, t - 1);
        if (!IsDurationRun(date, "YMD")) return false;
        if (t < 0) return date.Length != 0;
        string time = value.Substring(t + 1);
        return time.Length != 0 && IsDurationRun(time, "HMS");
    }

    /// <summary>Checks that the elements use consecutive units of the order, such as Y then M then D.</summary>
    private static bool IsDurationRun(string part, string order) {
        int i = 0, expected = -1;
        while (i < part.Length) {
            if (!TryDurationElement(part, ref i, out char unit)) return false;
            int position = order.IndexOf(unit);
            if (position < 0 || (expected >= 0 && position != expected)) return false;
            expected = position + 1;
        }
        return true;
    }

    private static bool TryDurationElement(string value, ref int i, out char unit) {
        unit = '\0';
        int start = i;
        while (i < value.Length && IsDigit(value[i])) i++;
        if (i == start || i >= value.Length) return false;
        unit = value[i++];
        return true;
    }

    #endregion

    #region Internet names and addresses

    private static bool IsEmail(string value, bool international) {
        int at = value.LastIndexOf('@');
        if (at <= 0 || at == value.Length - 1) return false;
        string local = value.Substring(0, at);
        string domain = value.Substring(at + 1);
        if (local.Length >= 2 && local[0] == '"' && local[local.Length - 1] == '"') {
            for (int i = 1; i < local.Length - 1; i++) {
                char c = local[i];
                if (c == '\\') c = local[++i];
                else if (c == '"') return false;
                if (c < 0x20 || (c > 0x7E && !international)) return false;
            }
        } else {
            if (local[0] == '.' || local[local.Length - 1] == '.' || local.IndexOf("..", StringComparison.Ordinal) >= 0) return false;
            foreach (char c in local) {
                if (c > 0x7F) {
                    if (!international) return false;
                    continue;
                }
                if (!(IsAsciiLetterOrDigit(c) || "!#$%&'*+-/=?^_`{|}~.".IndexOf(c) >= 0)) return false;
            }
        }
        if (domain.Length > 2 && domain[0] == '[' && domain[domain.Length - 1] == ']') {
            string literal = domain.Substring(1, domain.Length - 2);
            return literal.StartsWith("IPv6:", StringComparison.OrdinalIgnoreCase) ? IsIPv6(literal.Substring(5)) : IsIPv4(literal);
        }
        return international ? JsonSchemaHostnames.IsIdnHostname(domain) : JsonSchemaHostnames.IsHostname(domain);
    }

    private static bool IsIPv4(string value) {
        string[] parts = value.Split('.');
        if (parts.Length != 4) return false;
        foreach (string part in parts) {
            if (part.Length == 0 || part.Length > 3 || (part.Length > 1 && part[0] == '0')) return false;
            if (!TryDigits(part, 0, part.Length, out int octet) || octet > 255) return false;
        }
        return true;
    }

    private static bool IsIPv6(string value) {
        if (value.Length < 2 || value.Length > 45) return false;
        int dot = value.IndexOf('.');
        foreach (char c in value) {
            if (!(IsHexDigit(c) || c == ':' || c == '.')) return false;
        }
        if (dot >= 0 && !IsIPv4(value.Substring(value.LastIndexOf(':') + 1))) return false;
        return IPAddress.TryParse(value, out IPAddress? address) && address.AddressFamily == AddressFamily.InterNetworkV6;
    }

    #endregion

    #region URIs and pointers

    private static bool IsUri(string value, bool absolute, bool international) {
        if (value.Length == 0) return !absolute;
        bool hasScheme = JsonSchemaUri.HasScheme(value);
        if (absolute && !hasScheme) return false;
        int i = 0;
        if (hasScheme) i = value.IndexOf(':') + 1;
        int hash = value.IndexOf('#');
        bool inAuthority = value.Length >= i + 2 && value[i] == '/' && value[i + 1] == '/';
        int authorityEnd = inAuthority ? IndexOfAny(value, i + 2, '/', '?', '#') : -1;
        for (int p = 0; p < value.Length; p++) {
            char c = value[p];
            if (c == '%') {
                if (p + 2 >= value.Length || !IsHexDigit(value[p + 1]) || !IsHexDigit(value[p + 2])) return false;
                p += 2;
                continue;
            }
            if (c > 0x7F) {
                if (!international) return false;
                continue;
            }
            if (IsAsciiLetterOrDigit(c) || "-._~!$&'()*+,;=:@/?".IndexOf(c) >= 0) continue;
            if (c == '#' && p == hash) continue;
            // IP literals in the authority.
            if ((c == '[' || c == ']') && inAuthority && p < authorityEnd) continue;
            return false;
        }
        if (!hasScheme) {
            // A relative reference cannot start with a path segment containing ':'.
            int end = IndexOfAny(value, 0, '/', '?', '#');
            if (value.Substring(0, end).IndexOf(':') >= 0) return false;
        }
        if (inAuthority) {
            string authority = value.Substring(i + 2, authorityEnd - i - 2);
            int at = authority.IndexOf('@');
            if (at >= 0 && authority.IndexOf('@', at + 1) >= 0) return false;
            if (at >= 0 && (authority.IndexOf('[') < at && authority.IndexOf('[') >= 0 || authority.IndexOf(']') < at && authority.IndexOf(']') >= 0)) return false;
            string host = at >= 0 ? authority.Substring(at + 1) : authority;
            if (host.StartsWith("[", StringComparison.Ordinal)) {
                int close = host.IndexOf(']');
                if (close < 0) return false;
                string literal = host.Substring(1, close - 1);
                bool future = literal.Length > 1 && (literal[0] == 'v' || literal[0] == 'V');
                if (!future && !IsIPv6(literal)) return false;
                string port = host.Substring(close + 1);
                if (port.Length != 0 && (port[0] != ':' || !AllDigits(port, 1))) return false;
            } else {
                if (host.IndexOf('[') >= 0 || host.IndexOf(']') >= 0) return false;
                int colon = host.IndexOf(':');
                if (colon >= 0 && !AllDigits(host, colon + 1)) return false;
            }
        }
        return true;
    }

    /// <summary>RFC 6570 level 4 syntax: literals and <c>{op var:prefix,var*}</c> expressions.</summary>
    private static bool IsUriTemplate(string value) {
        int i = 0;
        while (i < value.Length) {
            char c = value[i];
            if (c == '{') {
                int close = value.IndexOf('}', i + 1);
                if (close < 0 || !IsUriTemplateExpression(value.Substring(i + 1, close - i - 1))) return false;
                i = close + 1;
                continue;
            }
            if (c == '%') {
                if (i + 2 >= value.Length || !IsHexDigit(value[i + 1]) || !IsHexDigit(value[i + 2])) return false;
                i += 3;
                continue;
            }
            if (c <= 0x20 || c == 0x7F || "\"<>\\^`|}".IndexOf(c) >= 0) return false;
            i++;
        }
        return true;
    }

    private static bool IsUriTemplateExpression(string expression) {
        if (expression.Length == 0) return false;
        int i = "+#./;?&".IndexOf(expression[0]) >= 0 ? 1 : 0;
        string[] specs = expression.Substring(i).Split(',');
        foreach (string spec in specs) {
            int end = spec.Length;
            if (spec.EndsWith("*", StringComparison.Ordinal)) {
                end--;
            } else {
                int colon = spec.IndexOf(':');
                if (colon >= 0) {
                    string length = spec.Substring(colon + 1);
                    if (length.Length is 0 or > 4 || length[0] == '0' || !AllDigits(length, 0)) return false;
                    end = colon;
                }
            }
            if (!IsUriTemplateVariable(spec.Substring(0, end))) return false;
        }
        return true;
    }

    private static bool IsUriTemplateVariable(string name) {
        if (name.Length == 0 || name[0] == '.' || name[name.Length - 1] == '.' || name.IndexOf("..", StringComparison.Ordinal) >= 0) return false;
        for (int i = 0; i < name.Length; i++) {
            char c = name[i];
            if (c == '%') {
                if (i + 2 >= name.Length || !IsHexDigit(name[i + 1]) || !IsHexDigit(name[i + 2])) return false;
                i += 2;
            } else if (!(IsAsciiLetterOrDigit(c) || c == '_' || c == '.')) {
                return false;
            }
        }
        return true;
    }

    private static bool IsUuid(string value) {
        if (value.Length != 36) return false;
        for (int i = 0; i < value.Length; i++) {
            bool dash = i is 8 or 13 or 18 or 23;
            if (dash ? value[i] != '-' : !IsHexDigit(value[i])) return false;
        }
        return true;
    }

    private static bool IsRegex(string value, TimeSpan timeout) {
        try {
            _ = JsonSchemaRegex.Create(value, timeout);
            return true;
        } catch (ArgumentException) {
            return false;
        }
    }

    private static bool IsJsonPointer(string value) => JsonPointerPath.TryParse(value, out _);

    private static bool IsRelativeJsonPointer(string value) {
        int i = 0;
        while (i < value.Length && IsDigit(value[i])) i++;
        if (i == 0 || (i > 1 && value[0] == '0')) return false;
        if (i < value.Length && (value[i] is '+' or '-')) {
            int start = ++i;
            while (i < value.Length && IsDigit(value[i])) i++;
            if (i == start || (i - start > 1 && value[start] == '0') || (i - start == 1 && value[start] == '0')) return false;
        }
        string rest = value.Substring(i);
        return rest == "#" || IsJsonPointer(rest);
    }

    #endregion

    private static int IndexOfAny(string value, int start, char a, char b, char c) {
        for (int i = start; i < value.Length; i++) {
            if (value[i] == a || value[i] == b || value[i] == c) return i;
        }
        return value.Length;
    }

    private static bool AllDigits(string value, int start) {
        for (int i = start; i < value.Length; i++) {
            if (!IsDigit(value[i])) return false;
        }
        return true;
    }

    private static bool TryDigits(string value, int start, int length, out int number) {
        number = 0;
        if (start + length > value.Length) return false;
        for (int i = start; i < start + length; i++) {
            if (!IsDigit(value[i])) return false;
            number = number * 10 + (value[i] - '0');
        }
        return true;
    }

    private static bool IsDigit(char c) => c >= '0' && c <= '9';

    private static bool IsHexDigit(char c) => IsDigit(c) || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');

    private static bool IsAsciiLetterOrDigit(char c) => IsDigit(c) || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
}
