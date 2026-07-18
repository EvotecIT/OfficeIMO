using OfficeIMO.Email;
using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

/// <summary>Outlook Rule FAI envelope with its client-owned opaque rule stream.</summary>
public sealed class EmailStoreRuleOrganizer {
    private const string ExpectedMessageClass = "IPM.RuleOrganizer";
    private const string ExpectedSubject = "Outlook Rules Organizer";

    internal EmailStoreRuleOrganizer(EmailDocument document) {
        MessageClass = document.MessageClass;
        Subject = document.Subject;
        RuleData = Copy(document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.RwRulesStream));
        if (RuleData != null) {
            using (SHA256 sha256 = SHA256.Create()) FingerprintSha256 = ToHex(sha256.ComputeHash(RuleData));
        }
    }

    /// <summary>Actual message class.</summary>
    public string? MessageClass { get; }
    /// <summary>Actual subject.</summary>
    public string? Subject { get; }
    /// <summary>Exact PidTagRwRulesStream bytes. Microsoft defines these client-owned bytes as opaque.</summary>
    public byte[]? RuleData { get; }
    /// <summary>Rule-stream byte length.</summary>
    public int RuleDataLength => RuleData?.Length ?? 0;
    /// <summary>Uppercase SHA-256 fingerprint for stable comparison without logging rule content.</summary>
    public string? FingerprintSha256 { get; }
    /// <summary>True when the Microsoft-defined message class and subject match.</summary>
    public bool IsProtocolEnvelopeValid =>
        string.Equals(MessageClass, ExpectedMessageClass, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(Subject, ExpectedSubject, StringComparison.Ordinal);

    private static byte[]? Copy(byte[]? value) => value == null ? null : (byte[])value.Clone();
    private static string ToHex(byte[] value) => string.Concat(value.Select(item => item.ToString("X2", CultureInfo.InvariantCulture)));
}
