using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace OfficeIMO.Entities;

/// <summary>Mailbox syntax and lengths from <see href="https://www.rfc-editor.org/rfc/rfc5321">RFC 5321</see>,
/// with <see href="https://www.rfc-editor.org/rfc/rfc5322">RFC 5322</see> syntax and <see href="https://www.rfc-editor.org/rfc/rfc6531">SMTPUTF8</see>.</summary>
internal sealed class EmailScheme : EntityScheme {
    private static readonly IdnMapping Idn = new() { UseStd3AsciiRules = true };

    public override string Id => "EMAIL";
    public override EntityKind Kind => EntityKind.Email;
    public override string DisplayName => "Email address";
    public override string Issuer => "Internet Engineering Task Force";
    public override Uri Specification { get; } = new("https://www.rfc-editor.org/rfc/rfc5321");
    public override int MinLength => 3;
    public override int MaxLength => 254;

    public override EntityValidationResult Validate(string? value) =>
        ValidateMailbox(value, allowLiteral: true, out _);

    internal EntityValidationResult ValidateMailbox(string? value, bool allowLiteral, out string? key) {
        key = null;
        if (string.IsNullOrWhiteSpace(value)) return new(this, EntityValidationStatus.Empty);
        string address = value!.Trim();
        if (address.Length < MinLength || address.Length > MaxLength)
            return new(this, EntityValidationStatus.InvalidLength);

        int at = FindSeparator(address);
        if (at <= 0 || at == address.Length - 1 || !IsLocalPart(address, at)) return new(this, EntityValidationStatus.InvalidFormat);
        string local = address.Substring(0, at);
        int localLength = Encoding.UTF8.GetByteCount(local);
        if (localLength > 64 || Encoding.UTF8.GetByteCount(address) > 254) return new(this, EntityValidationStatus.InvalidLength);

        string domain = address.Substring(at + 1).ToLowerInvariant();
        string? comparisonDomain = domain[0] == '[' ? (allowLiteral ? AddressLiteral(domain) : null) : AsciiDomain(domain);
        if (comparisonDomain == null) return new(this, EntityValidationStatus.InvalidFormat);
        if (localLength + 1 + comparisonDomain.Length > 254) return new(this, EntityValidationStatus.InvalidLength);

        key = local + "@" + comparisonDomain;
        return new EntityValidationResult(this, EntityValidationStatus.Valid, local + "@" + domain);
    }

    public override EntityDetails? Decode(string normalized) {
        EntityValidationResult result = ValidateMailbox(normalized, allowLiteral: true, out string? key);
        return result.IsValid ? new EntityDetails(properties: new Dictionary<string, string> { ["key"] = key! }) : null;
    }

    public override string MaskPartial(string normalized) {
        string? value = Normalize(normalized);
        if (value == null) return base.MaskPartial(normalized);
        int at = FindSeparator(value);
        int first = value[0] == '"' ? 1 : 0;
        if (first >= at || value[first] is '"' or '\\') return "•••" + value.Substring(at);
        int count = char.IsHighSurrogate(value[first]) ? 2 : 1;
        return value.Substring(first, count) + "•••" + value.Substring(at);
    }

    internal static bool IsAtom(char character) =>
        character is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9' ||
        "!#$%&'*+-/=?^_`{|}~".IndexOf(character) >= 0 ||
        character > 127 && !char.IsWhiteSpace(character) && !char.IsControl(character);

    internal static bool IsDomainCharacter(char character) =>
        char.IsLetterOrDigit(character) || character is '-' or '.' or '\u3002' or '\uFF0E' or '\uFF61' ||
        char.IsSurrogate(character) || CharUnicodeInfo.GetUnicodeCategory(character) is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark;

    private static int FindSeparator(string address) {
        if (address[0] != '"') return address.IndexOf('@');
        for (int index = 1; index < address.Length; index++) {
            if (address[index] == '\\') index++;
            else if (address[index] == '"') return index + 1 < address.Length && address[index + 1] == '@' ? index + 1 : -1;
        }
        return -1;
    }

    private static bool IsLocalPart(string value, int end) {
        bool quoted = value[0] == '"';
        if (quoted && (end < 2 || value[end - 1] != '"')) return false;
        bool needsAtom = true;
        for (int index = quoted ? 1 : 0; index < (quoted ? end - 1 : end); index++) {
            char character = value[index];
            if (char.IsSurrogate(character)) {
                if (index + 1 >= end || !char.IsSurrogatePair(value, index)) return false;
                index++;
                needsAtom = false;
            } else if (quoted && character == '\\') {
                if (++index >= end - 1 || value[index] < 32 || value[index] > 126) return false;
            } else if (quoted) {
                if (character == '"' || char.IsControl(character)) return false;
            } else if (character == '.') {
                if (needsAtom) return false;
                needsAtom = true;
            } else {
                if (!IsAtom(character)) return false;
                needsAtom = false;
            }
        }
        return quoted || !needsAtom;
    }

    private static string? AsciiDomain(string domain) {
        if (!IsDomainCharacter(domain[0]) || domain[domain.Length - 1] == '.') return null;
        foreach (char character in domain) {
            if (!IsDomainCharacter(character)) return null;
        }
        try {
            string ascii = Idn.GetAscii(domain).ToLowerInvariant();
            if (ascii.Length > 255 || ascii[ascii.Length - 1] == '.') return null;
            foreach (string label in ascii.Split('.')) {
                if (label.Length is < 1 or > 63 || label[0] == '-' || label[label.Length - 1] == '-') return null;
            }
            return ascii;
        } catch (ArgumentException) {
            return null;
        }
    }

    private static string? AddressLiteral(string domain) {
        if (domain.Length < 3 || domain[domain.Length - 1] != ']') return null;
        bool ipv6 = domain.StartsWith("[ipv6:", StringComparison.Ordinal);
        string address = ipv6 ? domain.Substring(6, domain.Length - 7) : domain.Substring(1, domain.Length - 2);
        if (address.IndexOf('%') >= 0) return null;
        if (!ipv6) {
            string[] parts = address.Split('.');
            if (parts.Length != 4) return null;
            foreach (string part in parts) {
                if (part.Length is < 1 or > 3) return null;
                foreach (char digit in part) {
                    if (digit is < '0' or > '9') return null;
                }
                if (int.Parse(part, CultureInfo.InvariantCulture) > 255) return null;
            }
            return domain;
        }
        return IPAddress.TryParse(address, out IPAddress? parsed) && parsed!.AddressFamily == AddressFamily.InterNetworkV6 ? domain : null;
    }
}
