using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace OfficeIMO.Provenance;

/// <summary>
/// Provides the dependency-free HTML manifest-association baseline used by Core. The OfficeIMO.Html
/// package remains the full DOM owner and additionally inspects embedded image resources.
/// </summary>
internal static class OfficeProvenanceHtml {
    private const string SyntheticBaseUri = "https://officeimo.invalid/__officeimo__/";

    internal static void Inspect(
        byte[] data,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext context) {
        string html = DecodeHtml(data, options.CancellationToken);
        var associations = new List<ManifestAssociation>();
        string? baseReference = null;
        bool inHead = false;
        bool headSeen = false;
        bool headFinished = false;
        bool bodyStarted = false;
        int elementCount = 0;
        int index = 0;

        while (index < html.Length) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int opening = html.IndexOf('<', index);
            if (opening < 0) break;
            if (StartsWith(html, opening, "<!--")) {
                int commentEnd = html.IndexOf("-->", opening + 4, StringComparison.Ordinal);
                if (commentEnd < 0) break;
                index = commentEnd + 3;
                continue;
            }

            if (!TryReadTag(html, opening, out HtmlTag tag)) {
                index = tag.End > opening ? tag.End : opening + 1;
                continue;
            }
            index = tag.End;
            if (tag.IsClosing) {
                if (tag.Name.Equals("head", StringComparison.OrdinalIgnoreCase)) {
                    inHead = false;
                    headFinished = true;
                }
                continue;
            }
            if (++elementCount > options.MaxContainerEntries) {
                throw new InvalidDataException("The HTML document exceeds the configured container-entry limit.");
            }

            if (tag.Name.Equals("head", StringComparison.OrdinalIgnoreCase)) {
                headSeen = true;
                inHead = true;
                continue;
            }
            if (tag.Name.Equals("body", StringComparison.OrdinalIgnoreCase)) {
                bodyStarted = true;
                inHead = false;
                headFinished = true;
                continue;
            }

            bool isHeadAssociation = inHead || (!headSeen && !headFinished && !bodyStarted);
            if (isHeadAssociation && baseReference == null &&
                tag.Name.Equals("base", StringComparison.OrdinalIgnoreCase) &&
                tag.Attributes.TryGetValue("href", out string? candidateBase)) {
                baseReference = candidateBase;
            }

            if (isHeadAssociation && tag.Name.Equals("link", StringComparison.OrdinalIgnoreCase) &&
                HasRelationship(tag.Attributes, "c2pa-manifest") &&
                tag.Attributes.TryGetValue("href", out string? href)) {
                associations.Add(ManifestAssociation.External(href));
            }

            if (!IsRawTextElement(tag.Name)) continue;
            string closingToken = "</" + tag.Name;
            int closeStart = FindClosingTag(html, index, closingToken);
            int contentEnd = closeStart < 0 ? html.Length : closeStart;
            if (tag.Name.Equals("script", StringComparison.OrdinalIgnoreCase) &&
                isHeadAssociation && tag.Attributes.TryGetValue("type", out string? type) &&
                TrimAsciiWhitespace(type).Equals("application/c2pa", StringComparison.OrdinalIgnoreCase)) {
                associations.Add(ManifestAssociation.Embedded(html.Substring(index, contentEnd - index)));
            }
            if (closeStart >= 0) {
                int closeEnd = html.IndexOf('>', closeStart + closingToken.Length);
                index = closeEnd < 0 ? html.Length : closeEnd + 1;
            } else {
                index = html.Length;
            }
        }

        bool singleAssociation = associations.Count == 1;
        if (!singleAssociation && associations.Count > 1) {
            context.Diagnostics.Add("HTML: manifest.html.multipleManifests: the HTML head contains multiple C2PA manifest associations.");
        }

        int carrierIndex = 0;
        foreach (ManifestAssociation association in associations) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (association.IsExternal) {
                string reference = NormalizeReference(association.Value);
                bool safeReference = TryValidateReference(reference, out Uri? parsed);
                bool valid = singleAssociation && safeReference;
                var evidence = new OfficeProvenanceEvidence(
                    OfficeProvenanceCarrierKind.C2paExternalManifest,
                    $"HTML/link[rel=c2pa-manifest][{carrierIndex++}]",
                    valid,
                    value: valid && parsed!.IsAbsoluteUri ? parsed.AbsoluteUri : reference);
                context.Add(evidence);
                if (valid) {
                    string resolved = ResolveReference(reference, baseReference);
                    if (!string.Equals(resolved, evidence.Value, StringComparison.Ordinal)) {
                        context.AddResolvedExternalManifestReference(evidence, resolved);
                    }
                }
                continue;
            }

            byte[] manifest = DecodeManifest(association.Value, options, context);
            bool embeddedValid = singleAssociation && manifest.Length != 0 &&
                OfficeC2paManifestStore.IsValid(
                    manifest,
                    0,
                    manifest.Length,
                    options.MaxManifestBytes,
                    options.MaxContainerEntries,
                    out _);
            context.Add(new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"HTML/script[type=application/c2pa][{carrierIndex++}]",
                embeddedValid,
                manifest.Length));
        }
    }

    private static byte[] DecodeManifest(
        string value,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext context) {
        string encoded = TrimAsciiWhitespace(value);
        const string prefix = "data:application/c2pa;base64,";
        if (encoded.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) encoded = encoded.Substring(prefix.Length);
        if (encoded.Length == 0 || encoded.Length > options.MaxManifestBytes * 2L || encoded.Length > int.MaxValue) {
            return Array.Empty<byte>();
        }
        long encodedCharacters = 0;
        for (int index = 0; index < encoded.Length; index++) {
            if ((index & 0xFFF) == 0) options.CancellationToken.ThrowIfCancellationRequested();
            if (!char.IsWhiteSpace(encoded[index])) encodedCharacters++;
        }
        long maximumEncodedCharacters = ((options.MaxManifestBytes + 2L) / 3L) * 4L;
        if (encodedCharacters > maximumEncodedCharacters) return Array.Empty<byte>();
        try {
            options.CancellationToken.ThrowIfCancellationRequested();
            byte[] manifest = Convert.FromBase64String(encoded);
            options.CancellationToken.ThrowIfCancellationRequested();
            if (manifest.LongLength > options.MaxManifestBytes) return Array.Empty<byte>();
            context.ReserveExpandedBytes(
                manifest.LongLength,
                "HTML provenance payloads exceed the configured expanded-container limit.");
            return manifest;
        } catch (FormatException) {
            return Array.Empty<byte>();
        }
    }

    private static string ResolveReference(string reference, string? baseReference) {
        if (string.IsNullOrWhiteSpace(baseReference)) return reference;
        string normalizedBase = NormalizeReference(baseReference!);
        if (!Uri.TryCreate(normalizedBase, UriKind.RelativeOrAbsolute, out Uri? parsedBase)) return reference;
        if (parsedBase.IsAbsoluteUri) {
            return Uri.TryCreate(parsedBase, reference, out Uri? absolute) ? absolute.AbsoluteUri : reference;
        }
        if (!Uri.TryCreate(new Uri(SyntheticBaseUri), parsedBase, out Uri? syntheticBase) ||
            !Uri.TryCreate(syntheticBase, reference, out Uri? resolved)) return reference;
        if (!resolved.AbsoluteUri.StartsWith(SyntheticBaseUri, StringComparison.Ordinal)) {
            return resolved.PathAndQuery + resolved.Fragment;
        }
        return resolved.AbsoluteUri.Substring(SyntheticBaseUri.Length);
    }

    private static bool TryValidateReference(string value, out Uri? uri) {
        uri = null;
        if (value.Length == 0 || !Uri.TryCreate(value, UriKind.RelativeOrAbsolute, out Uri? parsed)) return false;
        if (parsed.IsAbsoluteUri && parsed.Scheme != Uri.UriSchemeHttp && parsed.Scheme != Uri.UriSchemeHttps) return false;
        uri = parsed;
        return true;
    }

    private static string NormalizeReference(string? value) =>
        TrimAsciiWhitespace(value ?? string.Empty)
            .Replace("\t", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\r", string.Empty);

    private static bool HasRelationship(Dictionary<string, string> attributes, string relationship) {
        if (!attributes.TryGetValue("rel", out string? value)) return false;
        foreach (string token in value.Split(new[] { '\t', '\n', '\f', '\r', ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (token.Equals(relationship, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static bool TryReadTag(string html, int start, out HtmlTag tag) {
        tag = default;
        int index = start + 1;
        bool closing = index < html.Length && html[index] == '/';
        if (closing) index++;
        while (index < html.Length && IsAsciiWhitespace(html[index])) index++;
        int nameStart = index;
        while (index < html.Length && IsNameCharacter(html[index])) index++;
        if (index == nameStart) return false;
        string name = html.Substring(nameStart, index - nameStart);
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        while (index < html.Length) {
            while (index < html.Length && IsAsciiWhitespace(html[index])) index++;
            if (index >= html.Length) {
                tag = new HtmlTag(name, closing, attributes, html.Length);
                return false;
            }
            if (html[index] == '<') {
                tag = new HtmlTag(name, closing, attributes, index);
                return false;
            }
            if (html[index] == '>') {
                tag = new HtmlTag(name, closing, attributes, index + 1);
                return true;
            }
            if (html[index] == '/' && index + 1 < html.Length && html[index + 1] == '>') {
                tag = new HtmlTag(name, closing, attributes, index + 2);
                return true;
            }

            int attributeStart = index;
            while (index < html.Length && IsAttributeNameCharacter(html[index])) index++;
            if (index == attributeStart) {
                index++;
                continue;
            }
            string attributeName = html.Substring(attributeStart, index - attributeStart);
            while (index < html.Length && IsAsciiWhitespace(html[index])) index++;
            string attributeValue = string.Empty;
            if (index < html.Length && html[index] == '=') {
                index++;
                while (index < html.Length && IsAsciiWhitespace(html[index])) index++;
                if (index < html.Length && (html[index] == '\'' || html[index] == '"')) {
                    char quote = html[index++];
                    int valueStart = index;
                    while (index < html.Length && html[index] != quote) index++;
                    if (index >= html.Length) {
                        tag = new HtmlTag(name, closing, attributes, html.Length);
                        return false;
                    }
                    attributeValue = html.Substring(valueStart, index - valueStart);
                    index++;
                } else {
                    int valueStart = index;
                    while (index < html.Length && !IsAsciiWhitespace(html[index]) && html[index] != '>') index++;
                    attributeValue = html.Substring(valueStart, index - valueStart);
                }
            }
            if (!attributes.ContainsKey(attributeName)) {
                attributes.Add(attributeName, WebUtility.HtmlDecode(attributeValue));
            }
        }
        tag = new HtmlTag(name, closing, attributes, html.Length);
        return false;
    }

    private static int FindClosingTag(string html, int start, string token) {
        int search = start;
        while (search < html.Length) {
            int candidate = html.IndexOf(token, search, StringComparison.OrdinalIgnoreCase);
            if (candidate < 0) return -1;
            int boundary = candidate + token.Length;
            if (boundary >= html.Length || IsAsciiWhitespace(html[boundary]) || html[boundary] == '>') return candidate;
            search = boundary;
        }
        return -1;
    }

    internal static string DecodeHtml(byte[] data, System.Threading.CancellationToken cancellationToken) {
        Encoding encoding;
        int offset;
        if (data.Length >= 4 && data[0] == 0x00 && data[1] == 0x00 && data[2] == 0xFE && data[3] == 0xFF) {
            encoding = new UTF32Encoding(true, true);
            offset = 4;
        } else if (data.Length >= 4 && data[0] == 0xFF && data[1] == 0xFE && data[2] == 0x00 && data[3] == 0x00) {
            encoding = new UTF32Encoding(false, true);
            offset = 4;
        } else if (data.Length >= 2 && data[0] == 0xFE && data[1] == 0xFF) {
            encoding = Encoding.BigEndianUnicode;
            offset = 2;
        } else if (data.Length >= 2 && data[0] == 0xFF && data[1] == 0xFE) {
            encoding = Encoding.Unicode;
            offset = 2;
        } else {
            encoding = Encoding.UTF8;
            offset = data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF ? 3 : 0;
        }

        cancellationToken.ThrowIfCancellationRequested();
        Decoder decoder = encoding.GetDecoder();
        var output = new StringBuilder(Math.Min(data.Length - offset, 4096));
        var characters = new char[4096];
        int byteOffset = offset;
        bool completed;
        do {
            cancellationToken.ThrowIfCancellationRequested();
            decoder.Convert(
                data,
                byteOffset,
                data.Length - byteOffset,
                characters,
                0,
                characters.Length,
                flush: true,
                out int bytesUsed,
                out int charactersUsed,
                out completed);
            output.Append(characters, 0, charactersUsed);
            byteOffset += bytesUsed;
            if (!completed && bytesUsed == 0 && charactersUsed == 0) {
                throw new InvalidDataException("The HTML document could not be decoded incrementally.");
            }
        } while (!completed);
        cancellationToken.ThrowIfCancellationRequested();
        return output.ToString();
    }

    private static string TrimAsciiWhitespace(string value) {
        int start = 0;
        int end = value.Length;
        while (start < end && IsAsciiWhitespace(value[start])) start++;
        while (end > start && IsAsciiWhitespace(value[end - 1])) end--;
        return start == 0 && end == value.Length ? value : value.Substring(start, end - start);
    }

    private static bool StartsWith(string value, int offset, string candidate) =>
        offset >= 0 && offset <= value.Length - candidate.Length &&
        string.Compare(value, offset, candidate, 0, candidate.Length, StringComparison.Ordinal) == 0;

    private static bool IsAsciiWhitespace(char value) => value is '\t' or '\n' or '\f' or '\r' or ' ';
    private static bool IsRawTextElement(string name) =>
        name.Equals("script", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("style", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("title", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("textarea", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("xmp", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("iframe", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("noembed", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("noframes", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("plaintext", StringComparison.OrdinalIgnoreCase);
    private static bool IsNameCharacter(char value) => char.IsLetterOrDigit(value) || value is ':' or '-' or '_';
    private static bool IsAttributeNameCharacter(char value) =>
        !IsAsciiWhitespace(value) && value is not '=' and not '>' and not '/' and not '<' and not '\'' and not '"';

    private readonly struct HtmlTag {
        internal HtmlTag(string name, bool isClosing, Dictionary<string, string> attributes, int end) {
            Name = name;
            IsClosing = isClosing;
            Attributes = attributes;
            End = end;
        }

        internal string Name { get; }
        internal bool IsClosing { get; }
        internal Dictionary<string, string> Attributes { get; }
        internal int End { get; }
    }

    private readonly struct ManifestAssociation {
        private ManifestAssociation(string value, bool isExternal) {
            Value = value;
            IsExternal = isExternal;
        }

        internal string Value { get; }
        internal bool IsExternal { get; }
        internal static ManifestAssociation External(string value) => new ManifestAssociation(value, true);
        internal static ManifestAssociation Embedded(string value) => new ManifestAssociation(value, false);
    }
}
