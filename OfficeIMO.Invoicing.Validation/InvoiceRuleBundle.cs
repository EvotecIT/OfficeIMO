using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Validation;

/// <summary>Immutable, hash-verified authority artifacts loaded from local files. No network requests are made.</summary>
public sealed class InvoiceRuleBundle {
    /// <summary>Official KoSIT archive used for CII/UBL schemas, EN 16931 rules and XRechnung rules.</summary>
    public const string XRechnungDownloadUrl = "https://github.com/itplr-kosit/validator-configuration-xrechnung/releases/download/v2026-08-31/xrechnung-3.0.2-validator-configuration-2026-08-31.zip";
    /// <summary>SHA-256 of the pinned KoSIT archive.</summary>
    public const string XRechnungSha256 = "2530CD107C414511C5D0462EC10F886910395ABFCA820DB82E83D70BF01221A8";
    /// <summary>Immutable official Peppol source matching the published 3.0.21 artifact and pinned release hash.</summary>
    public const string PeppolDownloadUrl = "https://raw.githubusercontent.com/OpenPEPPOL/peppol-bis-invoice-3/806866bd2bd91d7e9623b68f08164e8fbe9e67a0/rules/sch/PEPPOL-EN16931-UBL.sch";
    /// <summary>SHA-256 of Peppol BIS Billing 3.0.21 Schematron.</summary>
    public const string PeppolSha256 = "62E5B67892F12755352D78B06F63229A02CC2ECCC748677C56EFBC8DBCB336E3";
    /// <summary>Immutable Apache-2.0 validator configuration containing the five Factur-X 1.09.2 profile schemas and rules.</summary>
    public const string FacturXDownloadUrl = "https://codeload.github.com/LandrixSoftware/validator-configuration-zugferd/zip/8399f5459df34b2af9cae0d613870649b3f17a58";
    /// <summary>Source commit of the pinned Factur-X/ZUGFeRD validator configuration.</summary>
    public const string FacturXSourceCommit = "8399f5459df34b2af9cae0d613870649b3f17a58";
    /// <summary>SHA-256 of the pinned Factur-X/ZUGFeRD validator configuration archive.</summary>
    public const string FacturXSha256 = "505D889923A3C58EA1EA59D65DA602C25656B918F55D6856B4BDA45B107B56D3";
    private readonly Dictionary<string, byte[]> _files;
    private readonly byte[]? _peppol;
    private readonly IReadOnlyDictionary<string, byte[]>? _facturXFiles;
    private InvoiceRuleBundle(Dictionary<string, byte[]> files, byte[]? peppol, IReadOnlyDictionary<string, byte[]>? facturXFiles) {
        _files = files;
        _peppol = peppol;
        _facturXFiles = facturXFiles;
    }

    /// <summary>Loads the pinned KoSIT archive and optional Peppol and Factur-X sources. Changed or corrupt artifacts are rejected.</summary>
    public static InvoiceRuleBundle Load(string xRechnungArchivePath, string? peppolSchematronPath = null, string? facturXArchivePath = null) {
        byte[] archive = ReadPinned(xRechnungArchivePath, XRechnungSha256, 8 * 1024 * 1024);
        byte[]? peppol = peppolSchematronPath == null ? null : ReadPinned(peppolSchematronPath, PeppolSha256, 2 * 1024 * 1024);
        Dictionary<string, byte[]> files = ReadArchive(archive, 32 * 1024 * 1024, static path => path);
        IReadOnlyDictionary<string, byte[]>? facturX = null;
        if (facturXArchivePath != null) {
            byte[] facturXArchive = ReadPinned(facturXArchivePath, FacturXSha256, 8 * 1024 * 1024);
            Dictionary<string, byte[]> extracted = ReadArchive(facturXArchive, 32 * 1024 * 1024, static path => {
                int separator = path.IndexOf('/');
                return separator < 0 ? null : path.Substring(separator + 1);
            });
            facturX = extracted.Where(item => item.Key.StartsWith("Schema/", StringComparison.Ordinal))
                .ToDictionary(item => item.Key, item => item.Value, StringComparer.Ordinal);
        }
        return new InvoiceRuleBundle(files, peppol, facturX);
    }

    /// <summary>True when the pinned optional Peppol rule source is loaded.</summary>
    public bool HasPeppolRules => _peppol != null;
    /// <summary>True when all five pinned Factur-X 1.09.2 profile schemas and rules are loaded.</summary>
    public bool HasFacturXRules => _facturXFiles != null;
    internal byte[] File(string path) {
        if (_files.TryGetValue(path, out byte[]? bytes)) return bytes;
        if (_facturXFiles != null && _facturXFiles.TryGetValue(path, out bytes)) return bytes;
        throw new InvalidDataException("Pinned bundle resource is missing: " + path);
    }
    internal byte[] Peppol => _peppol ?? throw new InvalidOperationException("Load the pinned Peppol Schematron source to validate Peppol BIS.");
    internal string Schema(InvoiceSpecificationRelease release, InvoiceProfile profile, InvoiceSyntax syntax, bool credit) {
        if (release == InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2) return FacturXProfile(profile).Schema;
        return syntax == InvoiceSyntax.Cii
            ? "resources/cii/16b/xsd/CrossIndustryInvoice_100pD16B.xsd"
            : "resources/ubl/2.1/xsd/maindoc/UBL-" + (credit ? "CreditNote" : "Invoice") + "-2.1.xsd";
    }
    internal IEnumerable<InvoiceRuleSource> Rules(InvoiceSpecificationRelease release, InvoiceProfile profile, InvoiceSyntax syntax) {
        if (release == InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2) {
            (string schema, string rules) = FacturXProfile(profile);
            _ = schema;
            IReadOnlyDictionary<string, byte[]> resources = _facturXFiles ?? throw new InvalidOperationException("Load the pinned Factur-X validation archive.");
            yield return new InvoiceRuleSource("Factur-X-1.09.2-" + profile, resources[rules], false, resources, rules);
            yield break;
        }
        string syntaxName = syntax == InvoiceSyntax.Cii ? "CII" : "UBL";
        string prefix = syntax == InvoiceSyntax.Cii ? "resources/cii/16b/xsl/" : "resources/ubl/2.1/xsl/";
        yield return new InvoiceRuleSource("EN16931-1.3.16-" + syntaxName, File(prefix + "EN16931-" + syntaxName + "-validation.xsl"), false);
        if (release == InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31)
            yield return new InvoiceRuleSource("XRechnung-3.0.2-2026-08-31-" + syntaxName, File("resources/xrechnung/3.0.2/xsl/XRechnung-" + syntaxName + "-validation.xsl"), false);
        if (release == InvoiceSpecificationRelease.PeppolBis_3_0_21) yield return new InvoiceRuleSource("Peppol-BIS-3.0.21-UBL", Peppol, true);
    }
    internal IReadOnlyDictionary<string, InvoiceDiagnosticSeverity> SeverityOverrides(InvoiceSpecificationRelease release, InvoiceSyntax syntax, bool credit) {
        var result = new Dictionary<string, InvoiceDiagnosticSeverity>(StringComparer.Ordinal);
        if (release != InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31) return result;
        XNamespace ns = "http://www.xoev.de/de/validator/framework/1/scenarios";
        using var input = new MemoryStream(File("scenarios.xml"), false);
        using var reader = XmlReader.Create(input, XmlSettings());
        XDocument document = XDocument.Load(reader);
        string scenarioName = syntax == InvoiceSyntax.Cii ? "EN16931 XRechnung (CII)" : "EN16931 XRechnung (UBL " + (credit ? "CreditNote" : "Invoice") + ")";
        XElement scenario = document.Root!.Elements(ns + "scenario").Single(s => (string?)s.Element(ns + "name") == scenarioName);
        foreach (XElement item in scenario.Element(ns + "createReport")!.Elements(ns + "customLevel"))
            result.Add(item.Value, ParseSeverity((string?)item.Attribute("level")));
        return result;
    }
    internal static InvoiceDiagnosticSeverity ParseSeverity(string? flag) => flag?.ToLowerInvariant() switch {
        "warning" => InvoiceDiagnosticSeverity.Warning,
        "information" or "info" => InvoiceDiagnosticSeverity.Information,
        _ => InvoiceDiagnosticSeverity.Error
    };
    internal static XmlReaderSettings XmlSettings() => new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16 * 1024 * 1024 };
    internal static byte[] ReadPinned(string path, string hash, long maximum) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = System.IO.File.OpenRead(path);
        if (stream.Length > maximum) throw new InvalidDataException("Pinned artifact exceeds its size limit: " + path);
        using var output = new MemoryStream();
        byte[] buffer = new byte[65536]; int count;
        while ((count = stream.Read(buffer, 0, buffer.Length)) != 0) {
            if (output.Length + count > maximum) throw new InvalidDataException("Pinned artifact exceeds its size limit: " + path);
            output.Write(buffer, 0, count);
        }
        byte[] bytes = output.ToArray();
        if (!string.Equals(Convert.ToHexString(SHA256.HashData(bytes)), hash, StringComparison.Ordinal)) throw new InvalidDataException("Pinned artifact SHA-256 does not match: " + path);
        return bytes;
    }

    private static Dictionary<string, byte[]> ReadArchive(byte[] archive, long maximumTotal, Func<string, string?> normalize) {
        var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using var input = new MemoryStream(archive, false);
        using var zip = new ZipArchive(input, ZipArchiveMode.Read);
        long total = 0;
        foreach (ZipArchiveEntry entry in zip.Entries) {
            if (entry.FullName.EndsWith('/')) continue;
            total += entry.Length;
            if (entry.Length > 8 * 1024 * 1024 || total > maximumTotal) throw new InvalidDataException("Rule bundle exceeds its uncompressed size limit.");
            string? path = normalize(entry.FullName);
            if (string.IsNullOrEmpty(path)) continue;
            using Stream stream = entry.Open();
            using var output = new MemoryStream();
            stream.CopyTo(output);
            files.Add(path, output.ToArray());
        }
        return files;
    }

    private static (string Schema, string Rules) FacturXProfile(InvoiceProfile profile) => profile switch {
        InvoiceProfile.Minimum => ("Schema/factur-x-1.09.2/0_Factur-X_1.09.2_MINIMUM/Factur-X_1.09.2_MINIMUM.xsd", "Schema/Factur-X-1.09.2_MINIMUM.xslt"),
        InvoiceProfile.BasicWithoutLines => ("Schema/factur-x-1.09.2/1_Factur-X_1.09.2_BASICWL/Factur-X_1.09.2_BASICWL.xsd", "Schema/Factur-X-1.09.2_BASICWL.xslt"),
        InvoiceProfile.Basic => ("Schema/factur-x-1.09.2/2_Factur-X_1.09.2_BASIC/Factur-X_1.09.2_BASIC.xsd", "Schema/Factur-X-1.09.2_BASIC.xslt"),
        InvoiceProfile.En16931 => ("Schema/factur-x-1.09.2/3_Factur-X_1.09.2_EN16931/Factur-X_1.09.2_EN16931.xsd", "Schema/Factur-X-1.09.2_EN16931.xslt"),
        InvoiceProfile.Extended => ("Schema/factur-x-1.09.2/4_Factur-X_1.09.2_EXTENDED/Factur-X_1.09.2_EXTENDED.xsd", "Schema/Factur-X-1.09.2_EXTENDED.xslt"),
        _ => throw new NotSupportedException("The selected profile is not part of Factur-X 1.09.2.")
    };
}

internal sealed record InvoiceRuleSource(string Name, byte[] Bytes, bool Compile,
    IReadOnlyDictionary<string, byte[]>? SupportingFiles = null, string? SourcePath = null);
