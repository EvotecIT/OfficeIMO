using System;

namespace OfficeIMO.Entities;

/// <summary>Web URL syntax based on <see href="https://url.spec.whatwg.org/">the WHATWG URL Standard</see>
/// and general URI components from <see href="https://www.rfc-editor.org/rfc/rfc3986">RFC 3986</see>.
/// Mailto internationalized domains follow <see href="https://www.rfc-editor.org/rfc/rfc6068">RFC 6068</see>.</summary>
internal sealed class UrlScheme : EntityScheme {
    public override string Id => "URL";
    public override EntityKind Kind => EntityKind.Url;
    public override string DisplayName => "URL";
    public override string Issuer => "WHATWG";
    public override Uri Specification { get; } = new("https://url.spec.whatwg.org/");
    public override int MinLength => 5;
    public override int MaxLength => 8192;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Public;

    public override EntityValidationResult Validate(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return new(this, EntityValidationStatus.Empty);
        string url = value!.Trim();
        if (url.Length < MinLength || url.Length > MaxLength) return new(this, EntityValidationStatus.InvalidLength);
        foreach (char character in url) {
            if (IsTerminator(character)) return new(this, EntityValidationStatus.InvalidCharacters);
        }

        if (url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) return ValidateMailto(url);
        bool webPrefix = url.StartsWith("www.", StringComparison.OrdinalIgnoreCase);
        int colon = webPrefix ? -1 : url.IndexOf(':');
        if (!webPrefix && (colon < 0 || colon + 2 >= url.Length || url[colon + 1] != '/' || url[colon + 2] != '/'))
            return new(this, EntityValidationStatus.InvalidFormat);
        string scheme = webPrefix ? "http" : url.Substring(0, colon).ToLowerInvariant();
        if (scheme is not ("http" or "https" or "ftp")) return new(this, EntityValidationStatus.InvalidFormat);

        if (!Uri.TryCreate(webPrefix ? "http://" + url : url, UriKind.Absolute, out Uri? parsed) ||
            parsed!.HostNameType is UriHostNameType.Unknown or UriHostNameType.IPv6) {
            return new(this, EntityValidationStatus.InvalidFormat);
        }

        int hostStart = webPrefix ? 0 : colon + 3;
        int authorityEnd = hostStart;
        while (authorityEnd < url.Length && url[authorityEnd] is not ('/' or '?' or '#')) authorityEnd++;
        for (int index = hostStart; index < authorityEnd; index++) {
            if (url[index] == '@') hostStart = index + 1;
        }
        int hostEnd = hostStart;
        while (hostEnd < authorityEnd && url[hostEnd] != ':') hostEnd++;

        string host = url.Substring(hostStart, hostEnd - hostStart);
        if (host.IndexOf('.') < 0 && !host.Equals("localhost", StringComparison.OrdinalIgnoreCase))
            return new(this, EntityValidationStatus.InvalidFormat);

        string writtenScheme = webPrefix ? string.Empty : scheme + url.Substring(colon, 3);
        int authorityStart = webPrefix ? 0 : colon + 3;
        string normalized = writtenScheme + url.Substring(authorityStart, hostStart - authorityStart) +
            host.ToLowerInvariant() + url.Substring(hostEnd);
        return new(this, EntityValidationStatus.Valid, normalized);
    }

    private EntityValidationResult ValidateMailto(string url) {
        int end = url.IndexOf('?', 7);
        if (end < 0) end = url.Length;
        string writtenMailbox = url.Substring(7, end - 7);
        EntityValidationResult result = ValidateMailtoMailbox(writtenMailbox, out bool encodedDomain);
        if (!result.IsValid) return new(this, EntityValidationStatus.InvalidFormat);

        string mailbox = result.Normalized!;
        int at = mailbox.IndexOf('@');
        string host = mailbox.Substring(at + 1);
        if (host.IndexOf('.') < 0 && host != "localhost") return new(this, EntityValidationStatus.InvalidFormat);
        if (encodedDomain) mailbox = mailbox.Substring(0, at + 1) + Uri.EscapeDataString(host);
        return new(this, EntityValidationStatus.Valid, "mailto:" + mailbox + url.Substring(end));
    }

    internal static EntityValidationResult ValidateMailtoMailbox(string mailbox, out bool encodedDomain) {
        int at = mailbox.LastIndexOf('@');
        encodedDomain = at >= 0 && mailbox.IndexOf('%', at + 1) >= 0;
        if (mailbox.Length == 0 || mailbox[0] == '"' || at < 1 || at == mailbox.Length - 1)
            return new EntityValidationResult(InternetSchemes.Email, EntityValidationStatus.InvalidFormat);
        string input = encodedDomain
            ? mailbox.Substring(0, at + 1) + Uri.UnescapeDataString(mailbox.Substring(at + 1))
            : mailbox;
        return InternetSchemes.EmailInstance.ValidateMailbox(input, allowLiteral: false, out _);
    }

    internal static bool IsTerminator(char character) => char.IsWhiteSpace(character) || char.IsControl(character) ||
        character is '<' or '>' or '"' or '\\' or '\u201C' or '\u201D' or '\u2018' or '\u2019';
}
