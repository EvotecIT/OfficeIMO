namespace OfficeIMO.Email;

/// <summary>Parsed RFC 6376 section 3.6.1 key record.</summary>
internal sealed class EmailDkimKeyRecord {
    private EmailDkimKeyRecord(string keyType, byte[] publicKey, IReadOnlyList<string>? hashes, IReadOnlyList<string> services,
        bool testing, bool strict) {
        KeyType = keyType;
        PublicKey = publicKey;
        Hashes = hashes;
        Services = services;
        Testing = testing;
        Strict = strict;
    }

    internal string KeyType { get; }

    internal byte[] PublicKey { get; }

    internal IReadOnlyList<string>? Hashes { get; }

    internal IReadOnlyList<string> Services { get; }

    internal bool Testing { get; }

    internal bool Strict { get; }

    internal bool IsRevoked => PublicKey.Length == 0;

    internal static EmailDkimKeyRecord? Parse(string record, out string error) {
        EmailDkimTagList? tags = EmailDkimTagList.Parse(record);
        if (tags == null) {
            error = "The key record is not a valid tag list.";
            return null;
        }

        string? version = tags["v"];
        if (version != null && (version != "DKIM1" || tags.Order[0] != "v")) {
            error = "The key record v= tag must be DKIM1 and appear first.";
            return null;
        }

        string? encodedKey = tags["p"];
        if (encodedKey == null) {
            error = "The key record has no p= tag.";
            return null;
        }

        byte[] key;
        try {
            key = Convert.FromBase64String(EmailDkimTagList.RemoveWhitespace(encodedKey));
        } catch (FormatException) {
            error = "The key record p= value is not valid base64.";
            return null;
        }

        string keyType = tags["k"] ?? "rsa";
        IReadOnlyList<string>? hashes = tags["h"] is string hashList ? SplitColonList(hashList) : null;
        IReadOnlyList<string> services = tags["s"] is string serviceList ? SplitColonList(serviceList) : new[] { "*" };
        IReadOnlyList<string> flags = tags["t"] is string flagList ? SplitColonList(flagList) : Array.Empty<string>();
        error = string.Empty;
        return new EmailDkimKeyRecord(keyType, key, hashes, services, flags.Contains("y"), flags.Contains("s"));
    }

    internal static string[] SplitColonList(string value) =>
        value.Split(':').Select(EmailDkimTagList.RemoveWhitespace).Where(item => item.Length > 0).ToArray();
}
