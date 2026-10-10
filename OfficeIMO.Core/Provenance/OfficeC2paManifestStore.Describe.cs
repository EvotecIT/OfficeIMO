using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Provenance;

internal static partial class OfficeC2paManifestStore {
    private const int MaximumDescribedActions = 64;
    private const int MaximumDescribedIngredients = 32;

    /// <summary>
    /// Reads what the active (last) manifest of a structurally valid store says. Never throws; returns null when the
    /// store or its claim cannot be read. Values are as written in the manifest and are not signature-verified.
    /// </summary>
    internal static OfficeC2paManifestSummary? TryDescribe(byte[] data, int offset, int length) {
        try {
            return Describe(data, offset, length);
        } catch (Exception error) when (error is FormatException or ArgumentException or IndexOutOfRangeException or OverflowException) {
            return null;
        }
    }

    private static OfficeC2paManifestSummary? Describe(byte[] data, int offset, int length) {
        if (data == null || offset < 0 || length <= 0 || offset > data.Length - length) return null;
        if (!TryOpenSuperbox(data, offset, length, out _, out string storeLabel, out int cursor, out int storeEnd) || storeLabel != "c2pa") return null;

        int manifestCount = 0, activeOffset = -1, activeLength = 0;
        var manifestIdentities = new List<string>();
        using var digest = System.Security.Cryptography.SHA256.Create();
        string? activeLabel = null;
        while (cursor < storeEnd) {
            if (!TryReadBox(data, cursor, storeEnd - cursor, out _, out ulong boxLength, out string type) || boxLength > int.MaxValue) return null;
            if (type == "jumb" && TryOpenSuperbox(data, cursor, (int)boxLength, out byte[] uuid, out string label, out _, out _) &&
                (SameUuid(uuid, StandardManifestUuid) || SameUuid(uuid, UpdateManifestUuid))) {
                manifestCount++;
                manifestIdentities.Add(Convert.ToBase64String(digest.ComputeHash(data, cursor, (int)boxLength)));
                activeOffset = cursor;
                activeLength = (int)boxLength;
                activeLabel = label;
            }
            cursor += (int)boxLength;
        }
        if (activeOffset < 0) return null;

        Dictionary<object, object?>? claim = null;
        var actions = new List<OfficeC2paAction>();
        bool declaresGenerativeAi = false;
        var ingredients = new List<string>();
        string? signedBy = null, issuer = null;
        var claimGenerators = new List<string>();

        TryOpenSuperbox(data, activeOffset, activeLength, out _, out _, out int child, out int manifestEnd);
        var assertionStores = new List<(int Offset, int End)>();
        while (child < manifestEnd) {
            if (!TryReadBox(data, child, manifestEnd - child, out _, out ulong childLength, out string childType) || childLength > int.MaxValue) break;
            if (childType == "jumb" && TryOpenSuperbox(data, child, (int)childLength, out byte[] uuid, out _, out int content, out int contentEnd)) {
                if (SameUuid(uuid, AssertionStoreUuid)) {
                    assertionStores.Add((content, contentEnd));
                } else if (SameUuid(uuid, ClaimUuid) && TryReadContent(data, content, contentEnd, out string claimType, out int claimOffset, out int claimLength) &&
                    claimType == "cbor" && OfficeCborReader.TryDecode(data, claimOffset, claimLength, out object? claimValue)) {
                    claim = claimValue as Dictionary<object, object?>;
                } else if (SameUuid(uuid, ClaimSignatureUuid) && TryReadContent(data, content, contentEnd, out string signatureType, out int signatureOffset, out int signatureLength) &&
                    signatureType == "cbor") {
                    ReadSigner(data, signatureOffset, signatureLength, out signedBy, out issuer);
                }
            }
            child += (int)childLength;
        }

        if (claim == null) return null;
        HashSet<string> claimedAssertions = ClaimedAssertions(claim, activeLabel!);
        var assertionBoxes = new List<(string Label, string Type, int Offset, int Length)>();
        foreach ((int content, int contentEnd) in assertionStores) CollectAssertions(data, content, contentEnd, claimedAssertions, assertionBoxes);
        string? generator = null;
        if (claim != null) {
            generator = Text(Get(claim, "claim_generator"));
            object? info = Get(claim, "claim_generator_info");
            if (info is Dictionary<object, object?> single) claimGenerators.Add(Agent(single) ?? string.Empty);
            if (info is List<object?> many) foreach (object? entry in many) claimGenerators.Add(Agent(entry) ?? string.Empty);
            if (claimGenerators.Count > 0 && !string.IsNullOrEmpty(claimGenerators[0])) generator = claimGenerators[0];
        }

        foreach ((string label, string type, int boxOffset, int boxLength) in assertionBoxes) {
            if (!claimedAssertions.Contains(label) || type != "cbor" || !OfficeCborReader.TryDecode(data, boxOffset, boxLength, out object? value) || value is not Dictionary<object, object?> assertion) continue;
            if (BaseLabel(label) == "c2pa.actions") {
                DescribeActions(assertion, label, actions, ref declaresGenerativeAi);
            } else if (BaseLabel(label) == "c2pa.ingredient" && ingredients.Count < MaximumDescribedIngredients) {
                string? title = Text(Get(assertion, "dc:title")) ?? Text(Get(assertion, "title"));
                if (!string.IsNullOrWhiteSpace(title)) ingredients.Add(title!);
            }
        }

        return new OfficeC2paManifestSummary(
            activeLabel,
            generator,
            claim == null ? null : Text(Get(claim, "dc:title")),
            claim == null ? null : Text(Get(claim, "dc:format")),
            actions,
            ingredients,
            signedBy,
            issuer,
            manifestCount,
            declaresGenerativeAi,
            manifestIdentities);
    }

    // A store may contain assertions which the active claim does not claim. Reading them as
    // origin evidence would let unrelated boxes change the displayed source classification.
    private static HashSet<string> ClaimedAssertions(Dictionary<object, object?> claim, string manifestLabel) {
        var labels = new HashSet<string>(StringComparer.Ordinal);
        string absolutePrefix = "self#jumbf=/c2pa/" + manifestLabel + "/c2pa.assertions/";
        const string relativePrefix = "self#jumbf=c2pa.assertions/";
        foreach (string field in new[] { "assertions", "created_assertions", "gathered_assertions" }) {
            if (Get(claim, field) is not List<object?> references) continue;
            foreach (object? reference in references) {
                if (reference is not Dictionary<object, object?> map || Get(map, "url") is not string url ||
                    Get(map, "hash") is not byte[] hash || hash.Length == 0) continue;
                string? label = url.StartsWith(relativePrefix, StringComparison.Ordinal) ? url.Substring(relativePrefix.Length)
                    : url.StartsWith(absolutePrefix, StringComparison.Ordinal) ? url.Substring(absolutePrefix.Length) : null;
                if (!string.IsNullOrEmpty(label) && label!.IndexOf('/') < 0 && label != "." && label != "..") labels.Add(label);
            }
        }
        return labels; // Hash membership is read here; integrity and signature verification are separate operations.
    }

    private static void CollectAssertions(byte[] data, int cursor, int end, HashSet<string> claimed, List<(string, string, int, int)> boxes) {
        while (cursor < end) {
            if (!TryReadBox(data, cursor, end - cursor, out _, out ulong length, out string type) || length > int.MaxValue) return;
            if (type == "jumb" && TryOpenSuperbox(data, cursor, (int)length, out _, out string label, out int content, out int contentEnd) &&
                claimed.Contains(label) && BaseLabel(label) is "c2pa.actions" or "c2pa.ingredient" &&
                TryReadContent(data, content, contentEnd, out string contentType, out int payloadOffset, out int payloadLength)) {
                boxes.Add((label, contentType, payloadOffset, payloadLength));
            }
            cursor += (int)length;
        }
    }

    /// <summary>Opens a JUMBF superbox: returns its description UUID and label, and the range of its child boxes.</summary>
    private static bool TryOpenSuperbox(byte[] data, int offset, int length, out byte[] uuid, out string label, out int contentOffset, out int end) {
        uuid = Array.Empty<byte>();
        label = string.Empty;
        contentOffset = 0;
        end = 0;
        if (!TryReadBox(data, offset, length, out int headerLength, out ulong declaredLength, out string type) || type != "jumb" || declaredLength > (ulong)length) return false;
        int descriptionOffset = offset + headerLength;
        int superboxEnd = offset + (int)declaredLength;
        if (!TryReadBox(data, descriptionOffset, superboxEnd - descriptionOffset, out int descriptionHeaderLength, out ulong descriptionLength, out string descriptionType) ||
            descriptionType != "jumd" || descriptionLength < (ulong)(descriptionHeaderLength + 17)) return false;
        int payload = descriptionOffset + descriptionHeaderLength;
        uuid = new byte[16];
        Buffer.BlockCopy(data, payload, uuid, 0, 16);
        int descriptionEnd = descriptionOffset + (int)descriptionLength;
        if (!TryReadDescriptionFields(data, payload + 16, descriptionEnd, out label)) return false;
        contentOffset = descriptionEnd;
        end = superboxEnd;
        return true;
    }

    /// <summary>Reads the first content box (cbor, json, …) of a superbox.</summary>
    private static bool TryReadContent(byte[] data, int offset, int end, out string type, out int payloadOffset, out int payloadLength) {
        payloadOffset = payloadLength = 0;
        if (!TryReadBox(data, offset, end - offset, out int headerLength, out ulong length, out type) || length > int.MaxValue) return false;
        payloadOffset = offset + headerLength;
        payloadLength = (int)length - headerLength;
        return payloadLength > 0;
    }

    /// <summary>Reads the signer from the COSE_Sign1 x5chain (header 33): the first certificate is the signing certificate.</summary>
    private static void ReadSigner(byte[] data, int offset, int length, out string? signedBy, out string? issuer) {
        signedBy = issuer = null;
        if (!OfficeCborReader.TryDecode(data, offset, length, out object? value)) return;
        if (value is OfficeCborTag tag) value = tag.Value;
        if (value is not List<object?> sign1 || sign1.Count < 2) return;
        object? chain = null;
        if (sign1[0] is byte[] protectedHeader && protectedHeader.Length > 0 &&
            OfficeCborReader.TryDecode(protectedHeader, 0, protectedHeader.Length, out object? decoded) && decoded is Dictionary<object, object?> protectedMap) {
            chain = Get(protectedMap, 33L);
        }
        if (chain == null && sign1[1] is Dictionary<object, object?> unprotectedMap) chain = Get(unprotectedMap, 33L) ?? Get(unprotectedMap, "x5chain");
        byte[]? certificate = chain as byte[] ?? (chain is List<object?> certificates && certificates.Count > 0 ? certificates[0] as byte[] : null);
        if (certificate == null) return;
        OfficeX509Name.TryRead(certificate, out signedBy, out issuer);
    }

    private static string BaseLabel(string label) {
        // Assertion labels carry a version and an optional instance suffix, for example c2pa.actions.v2 or c2pa.ingredient__1.
        int instance = label.IndexOf("__", StringComparison.Ordinal);
        string name = instance >= 0 ? label.Substring(0, instance) : label;
        int version = name.LastIndexOf(".v", StringComparison.Ordinal);
        if (version > 0 && version + 2 < name.Length && char.IsDigit(name[version + 2])) name = name.Substring(0, version);
        return name;
    }

    private static string? Agent(object? value) {
        if (value is string text) return Text(text);
        if (value is not Dictionary<object, object?> map) return null;
        string? name = Text(Get(map, "name"));
        string? version = Text(Get(map, "version"));
        return name == null ? null : version == null ? name : name + " " + version;
    }

    private static object? Get(Dictionary<object, object?> map, object key) => map.TryGetValue(key, out object? value) ? value : null;

    private static string? Text(object? value) {
        if (value is not string text) return null;
        text = text.Trim();
        if (text.Length == 0) return null;
        return text.Length > 300 ? text.Substring(0, 300) : text;
    }

    private static bool SameUuid(byte[] uuid, byte[] expected) {
        if (uuid.Length != expected.Length) return false;
        for (int index = 0; index < uuid.Length; index++) if (uuid[index] != expected[index]) return false;
        return true;
    }
}

/// <summary>Reads the subject and issuer names of a DER X.509 certificate without platform crypto (works on netstandard2.0 and WebAssembly).</summary>
internal static class OfficeX509Name {
    private static readonly byte[] CommonName = { 0x55, 0x04, 0x03 };
    private static readonly byte[] Organization = { 0x55, 0x04, 0x0A };

    internal static bool TryRead(byte[] certificate, out string? subject, out string? issuer) {
        subject = issuer = null;
        try {
            int position = 0;
            if (!Enter(certificate, ref position, certificate.Length, 0x30, out int certificateEnd)) return false;  // Certificate
            if (!Enter(certificate, ref position, certificateEnd, 0x30, out int tbsEnd)) return false;              // TBSCertificate
            if (position < tbsEnd && certificate[position] == 0xA0) Skip(certificate, ref position, tbsEnd);       // [0] version
            Skip(certificate, ref position, tbsEnd);                                                                // serialNumber
            Skip(certificate, ref position, tbsEnd);                                                                // signature algorithm
            issuer = ReadName(certificate, ref position, tbsEnd);
            Skip(certificate, ref position, tbsEnd);                                                                // validity
            subject = ReadName(certificate, ref position, tbsEnd);
            return subject != null || issuer != null;
        } catch (FormatException) {
            subject = issuer = null;
            return false;
        }
    }

    private static string? ReadName(byte[] data, ref int position, int limit) {
        int start = position;
        if (!Enter(data, ref position, limit, 0x30, out int nameEnd)) return null;
        string? organization = null, commonName = null;
        while (position < nameEnd) {
            if (!Enter(data, ref position, nameEnd, 0x31, out int setEnd)) break;
            while (position < setEnd) {
                if (!Enter(data, ref position, setEnd, 0x30, out int attributeEnd)) break;
                ReadHeader(data, ref position, attributeEnd, out byte oidTag, out int oidLength);
                bool isOrganization = oidTag == 0x06 && Matches(data, position, oidLength, Organization);
                bool isCommonName = oidTag == 0x06 && Matches(data, position, oidLength, CommonName);
                position += oidLength;
                ReadHeader(data, ref position, attributeEnd, out byte valueTag, out int valueLength);
                string? text = DecodeString(data, position, valueLength, valueTag);
                if (isOrganization) organization ??= text;
                if (isCommonName) commonName ??= text;
                position = attributeEnd;
            }
            position = setEnd;
        }
        position = nameEnd;
        if (start == position) return null;
        return organization ?? commonName;
    }

    private static bool Enter(byte[] data, ref int position, int limit, byte expectedTag, out int end) {
        ReadHeader(data, ref position, limit, out byte tag, out int length);
        end = position + length;
        return tag == expectedTag;
    }

    private static void Skip(byte[] data, ref int position, int limit) {
        ReadHeader(data, ref position, limit, out _, out int length);
        position += length;
    }

    private static void ReadHeader(byte[] data, ref int position, int limit, out byte tag, out int length) {
        if (position + 2 > limit) throw new FormatException("Truncated certificate.");
        tag = data[position++];
        int first = data[position++];
        if (first < 0x80) {
            length = first;
        } else {
            int bytes = first & 0x7F;
            if (bytes == 0 || bytes > 4 || position + bytes > limit) throw new FormatException("Unsupported certificate length.");
            length = 0;
            for (int index = 0; index < bytes; index++) length = (length << 8) | data[position++];
        }
        if (length < 0 || length > limit - position) throw new FormatException("Certificate field exceeds its container.");
    }

    private static bool Matches(byte[] data, int position, int length, byte[] expected) {
        if (length != expected.Length) return false;
        for (int index = 0; index < length; index++) if (data[position + index] != expected[index]) return false;
        return true;
    }

    private static string? DecodeString(byte[] data, int position, int length, byte tag) {
        string? text = tag switch {
            0x0C => Encoding.UTF8.GetString(data, position, length),                           // UTF8String
            0x13 or 0x16 or 0x14 => Encoding.ASCII.GetString(data, position, length),          // Printable, IA5, Teletex
            0x1E => Encoding.BigEndianUnicode.GetString(data, position, length),               // BMPString
            _ => null
        };
        text = text?.Trim();
        return string.IsNullOrEmpty(text) ? null : text!.Length > 200 ? text.Substring(0, 200) : text;
    }
}
