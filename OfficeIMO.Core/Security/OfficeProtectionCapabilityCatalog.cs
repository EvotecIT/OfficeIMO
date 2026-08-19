using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Security;

/// <summary>Machine-readable source of truth for protected-content coverage across OfficeIMO formats.</summary>
public sealed class OfficeProtectionCapabilityCatalog {
    private readonly IReadOnlyDictionary<string, OfficeProtectionCapability> _byId;

    /// <summary>Current protected-content capability contract.</summary>
    public static OfficeProtectionCapabilityCatalog Current { get; } = new OfficeProtectionCapabilityCatalog(
        "OfficeIMO.ProtectedContent", 1, CreateCurrentRows());

    /// <summary>Creates a protected-content capability catalog.</summary>
    public OfficeProtectionCapabilityCatalog(string id, int schemaVersion,
        IEnumerable<OfficeProtectionCapability> capabilities) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Catalog id cannot be empty.", nameof(id));
        if (schemaVersion <= 0) throw new ArgumentOutOfRangeException(nameof(schemaVersion));
        if (capabilities == null) throw new ArgumentNullException(nameof(capabilities));
        OfficeProtectionCapability[] rows = capabilities.ToArray();
        if (rows.Length == 0) throw new ArgumentException("A capability catalog must contain at least one row.", nameof(capabilities));
        string[] duplicates = rows.GroupBy(row => row.Id, StringComparer.Ordinal)
            .Where(group => group.Count() > 1).Select(group => group.Key).ToArray();
        if (duplicates.Length != 0) throw new ArgumentException("Capability ids must be unique: " + string.Join(", ", duplicates), nameof(capabilities));
        Id = id.Trim();
        SchemaVersion = schemaVersion;
        Capabilities = new ReadOnlyCollection<OfficeProtectionCapability>(rows);
        _byId = new ReadOnlyDictionary<string, OfficeProtectionCapability>(rows.ToDictionary(row => row.Id, StringComparer.Ordinal));
    }

    /// <summary>Stable catalog identifier.</summary>
    public string Id { get; }
    /// <summary>Schema version.</summary>
    public int SchemaVersion { get; }
    /// <summary>Capability rows in stable order.</summary>
    public IReadOnlyList<OfficeProtectionCapability> Capabilities { get; }
    /// <summary>Rows that contain at least one unsupported operation.</summary>
    public IReadOnlyList<OfficeProtectionCapability> IncompleteCapabilities => Capabilities.Where(row =>
        row.Inspect == OfficeProtectionCoverageState.NotSupported ||
        row.Open == OfficeProtectionCoverageState.NotSupported ||
        row.Create == OfficeProtectionCoverageState.NotSupported ||
        row.Validate == OfficeProtectionCoverageState.NotSupported ||
        row.Mutate == OfficeProtectionCoverageState.NotSupported ||
        row.Remove == OfficeProtectionCoverageState.NotSupported).ToArray();

    /// <summary>Gets a capability by exact stable id.</summary>
    public OfficeProtectionCapability Get(string id) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Capability id cannot be empty.", nameof(id));
        if (!_byId.TryGetValue(id.Trim(), out OfficeProtectionCapability? capability)) {
            throw new KeyNotFoundException($"Capability '{id}' is not present in catalog '{Id}'.");
        }
        return capability;
    }

    /// <summary>Serializes the catalog as deterministic JSON without requiring a JSON package.</summary>
    public string ToJson() {
        var output = new StringBuilder();
        output.Append("{\n  \"id\":\"").Append(EscapeJson(Id)).Append("\",\n  \"schemaVersion\":")
            .Append(SchemaVersion).Append(",\n  \"capabilities\":[\n");
        for (int index = 0; index < Capabilities.Count; index++) {
            OfficeProtectionCapability row = Capabilities[index];
            output.Append("    {\n")
                .Append("      \"id\":\"").Append(EscapeJson(row.Id)).Append("\",\n")
                .Append("      \"formatId\":\"").Append(EscapeJson(row.FormatId)).Append("\",\n")
                .Append("      \"packageId\":\"").Append(EscapeJson(row.PackageId)).Append("\",\n")
                .Append("      \"kind\":\"").Append(row.Kind).Append("\",\n")
                .Append("      \"inspect\":\"").Append(row.Inspect).Append("\",\n")
                .Append("      \"open\":\"").Append(row.Open).Append("\",\n")
                .Append("      \"create\":\"").Append(row.Create).Append("\",\n")
                .Append("      \"validate\":\"").Append(row.Validate).Append("\",\n")
                .Append("      \"mutate\":\"").Append(row.Mutate).Append("\",\n")
                .Append("      \"remove\":\"").Append(row.Remove).Append("\",\n")
                .Append("      \"api\":\"").Append(EscapeJson(row.Api)).Append("\",\n")
                .Append("      \"limitation\":\"").Append(EscapeJson(row.Limitation)).Append("\"\n    }");
            if (index + 1 < Capabilities.Count) output.Append(',');
            output.Append('\n');
        }
        return output.Append("  ]\n}").ToString();
    }

    /// <summary>Formats the catalog as a deterministic Markdown table.</summary>
    public string ToMarkdown() {
        var output = new StringBuilder();
        output.Append("# ").Append(Id).Append(" capability contract\n\nSchema version: ").Append(SchemaVersion)
            .Append("\n\n| Capability | Format | Owner | Kind | Inspect | Open | Create | Validate | Mutate | Remove | API | Limitation |\n")
            .Append("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |\n");
        foreach (OfficeProtectionCapability row in Capabilities) {
            output.Append("| ").Append(EscapeMarkdown(row.Id)).Append(" | ").Append(EscapeMarkdown(row.FormatId))
                .Append(" | ").Append(EscapeMarkdown(row.PackageId)).Append(" | ").Append(row.Kind)
                .Append(" | ").Append(row.Inspect).Append(" | ").Append(row.Open).Append(" | ").Append(row.Create)
                .Append(" | ").Append(row.Validate).Append(" | ").Append(row.Mutate).Append(" | ").Append(row.Remove)
                .Append(" | `").Append(EscapeMarkdown(row.Api)).Append("` | ").Append(EscapeMarkdown(row.Limitation)).Append(" |\n");
        }
        return output.ToString();
    }

    private static OfficeProtectionCapability[] CreateCurrentRows() => new[] {
        Row("ooxml-password", "DOCX/XLSX/PPTX", "OfficeIMO.Word / OfficeIMO.Excel / OfficeIMO.PowerPoint", OfficeProtectionKind.PasswordEncryption,
            S(), S(), S(), N(), S(), S(), "format load/save password APIs", "Password encryption is format-owned; no OfficeIMO.Security dependency is required."),
        Row("doc-password", "DOC", "OfficeIMO.Word", OfficeProtectionKind.PasswordEncryption,
            D(), NS(), NS(), N(), B(), NS(), "WordDocument.Load", "Legacy DOC encryption is detected only; encrypted content is not exposed."),
        Row("xls-password", "XLS", "OfficeIMO.Excel", OfficeProtectionKind.PasswordEncryption,
            S(), S(), NS(), N(), NS(), NS(), "ExcelDocument.Load", "Encrypted legacy XLS read is supported; encrypted legacy authoring is not."),
        Row("ppt-password", "PPT", "OfficeIMO.PowerPoint", OfficeProtectionKind.PasswordEncryption,
            S(), S(), S(), N(), S(), S(), "PowerPointPresentation.Load / Save", "Legacy PPT RC4 password protection is format-owned."),
        Row("pdf-password", "PDF", "OfficeIMO.Pdf", OfficeProtectionKind.PasswordEncryption,
            S(), S(), S(), N(), S(), S(), "PdfDocument.Security", "Typed PDF security options and operation results report encryption policy."),
        Row("odf-password", "ODT/ODS/ODP", "OfficeIMO.OpenDocument", OfficeProtectionKind.PasswordEncryption,
            S(), S(), S(), N(), S(), S(), "OdfLoadOptions.Password / OdfSaveOptions.Encryption", "AES-256-CBC with 10,000-10,000,000 PBKDF2-HMAC-SHA1 iterations, SHA-256 start key, and SHA-256/1K checksum; legacy Blowfish is rejected."),
        Row("epub-font-obfuscation", "EPUB", "OfficeIMO.Epub", OfficeProtectionKind.FontObfuscation,
            S(), S(), NS(), N(), N(), N(), "EpubDocument.Load", "IDPF and Adobe font obfuscation are removed in the read projection only when the package identity yields the required key; EPUB output is not mutated."),
        Row("smime-signature", "EML / RFC 5322 MIME", "OfficeIMO.Email + optional provider", OfficeProtectionKind.DigitalSignature,
            S(), S(), S(), S(), P(), N(), "EmailSmime.Sign / Verify", "Email owns MIME serialization; an explicit IOfficeSecurityProvider owns CMS and certificate policy."),
        Row("smime-signature-msg-tnef", "MSG/TNEF S/MIME", "OfficeIMO.Email + optional provider", OfficeProtectionKind.DigitalSignature,
            S(), S(), NS(), S(), P(), N(), "EmailSmime.Verify", "Protected MSG/TNEF content can be inspected and verified after parsing; outbound signing creates RFC 5322/MIME bytes, not MSG or TNEF containers."),
        Row("smime-envelope", "EML / RFC 5322 MIME", "OfficeIMO.Email + optional provider", OfficeProtectionKind.RecipientEncryption,
            S(), S(), S(), N(), P(), N(), "EmailSmime.Encrypt / Decrypt / SignAndEncrypt", "Certificate-recipient CMS encryption is provider-backed; protected source pass-through remains available without the provider."),
        Row("smime-envelope-msg-tnef", "MSG/TNEF S/MIME", "OfficeIMO.Email + optional provider", OfficeProtectionKind.RecipientEncryption,
            S(), S(), NS(), N(), P(), N(), "EmailSmime.Decrypt", "Protected MSG/TNEF content can be inspected and decrypted after parsing; outbound encryption creates RFC 5322/MIME bytes, not MSG or TNEF containers."),
        Row("opc-package-signature", "DOCX/XLSX/PPTX/Visio", "format package + optional provider", OfficeProtectionKind.DigitalSignature,
            S(), N(), S(), S(), B(), S(), "InspectPackageSignatures / SignPackageSignature / ValidatePackageSignatures", "Changed signed packages require explicit invalidation handling."),
        Row("vba-signature", "DOCM/XLSM/XLSB/PPTM families", "format package + optional provider", OfficeProtectionKind.DigitalSignature,
            S(), N(), S(), S(), B(), S(), "InspectVbaSignatures / SignVbaProject", "Managed legacy, agile, and V3 VBA signature profiles are corpus-bound."),
        Row("odf-package-signature", "ODT/ODS/ODP", "OfficeIMO.OpenDocument + optional provider", OfficeProtectionKind.DigitalSignature,
            S(), N(), S(), S(), B(), S(), "OdfDocument.SignPackage / ValidatePackageSignatures", "The bounded OfficeIMO XML manifest profile does not claim every producer-specific ODF signature profile."),
        Row("epub-package-signature", "EPUB", "OfficeIMO.Epub + optional provider", OfficeProtectionKind.DigitalSignature,
            S(), N(), S(), S(), B(), S(), "EpubDocument.SignPackage / ValidatePackageSignatures", "The bounded OfficeIMO XML manifest profile does not claim arbitrary EPUB DRM or signature profiles."),
        Row("onenote-encrypted-revision", "ONE", "OfficeIMO.OneNote", OfficeProtectionKind.PasswordEncryption,
            D(), NS(), NS(), N(), B(), NS(), "OneNoteSectionReader.Read", "An encrypted current revision or dependency fails closed; an older plaintext revision is never substituted."),
        Row("pst-password", "PST", "OfficeIMO.Email", OfficeProtectionKind.AccessDeterrence,
            S(), S(), NS(), N(), B(), NS(), "EmailStoreSession.IsPstPasswordProtected", "PST password protection is a checksum-based access deterrent, not cryptographic encryption."),
        Row("rtf-editing-restrictions", "RTF", "OfficeIMO.Rtf", OfficeProtectionKind.EditingRestriction,
            S(), N(), S(), N(), S(), S(), "RtfDocumentSettings.HasEditingProtection", "RTF protection flags restrict editing but do not encrypt content.")
    };

    private static OfficeProtectionCapability Row(string id, string format, string package, OfficeProtectionKind kind,
        OfficeProtectionCoverageState inspect, OfficeProtectionCoverageState open, OfficeProtectionCoverageState create,
        OfficeProtectionCoverageState validate, OfficeProtectionCoverageState mutate, OfficeProtectionCoverageState remove,
        string api, string limitation) => new OfficeProtectionCapability(id, format, package, kind, inspect, open,
        create, validate, mutate, remove, api, limitation);
    private static OfficeProtectionCoverageState S() => OfficeProtectionCoverageState.Supported;
    private static OfficeProtectionCoverageState D() => OfficeProtectionCoverageState.Detected;
    private static OfficeProtectionCoverageState P() => OfficeProtectionCoverageState.Preserved;
    private static OfficeProtectionCoverageState B() => OfficeProtectionCoverageState.Blocked;
    private static OfficeProtectionCoverageState NS() => OfficeProtectionCoverageState.NotSupported;
    private static OfficeProtectionCoverageState N() => OfficeProtectionCoverageState.NotApplicable;
    private static string EscapeJson(string value) {
        var escaped = new StringBuilder(value.Length + 8);
        foreach (char character in value) {
            switch (character) {
                case '\"': escaped.Append("\\\""); break;
                case '\\': escaped.Append("\\\\"); break;
                case '\b': escaped.Append("\\b"); break;
                case '\f': escaped.Append("\\f"); break;
                case '\n': escaped.Append("\\n"); break;
                case '\r': escaped.Append("\\r"); break;
                case '\t': escaped.Append("\\t"); break;
                default:
                    if (character < ' ') {
                        escaped.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                    } else {
                        escaped.Append(character);
                    }
                    break;
            }
        }
        return escaped.ToString();
    }
    private static string EscapeMarkdown(string value) => value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
}
