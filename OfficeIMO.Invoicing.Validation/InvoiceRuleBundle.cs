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
    private readonly Dictionary<string, byte[]> _files;
    private readonly byte[]? _peppol;
    private InvoiceRuleBundle(Dictionary<string, byte[]> files, byte[]? peppol) { _files = files; _peppol = peppol; }

    /// <summary>Loads the pinned KoSIT archive and optionally the pinned Peppol source. A changed or corrupt artifact is rejected.</summary>
    public static InvoiceRuleBundle Load(string xRechnungArchivePath, string? peppolSchematronPath = null) {
        byte[] archive = ReadPinned(xRechnungArchivePath, XRechnungSha256, 8 * 1024 * 1024);
        byte[]? peppol = peppolSchematronPath == null ? null : ReadPinned(peppolSchematronPath, PeppolSha256, 2 * 1024 * 1024);
        var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using var input = new MemoryStream(archive, false);
        using var zip = new ZipArchive(input, ZipArchiveMode.Read);
        long total = 0;
        foreach (ZipArchiveEntry entry in zip.Entries) {
            if (entry.FullName.EndsWith('/')) continue;
            total += entry.Length;
            if (entry.Length > 8 * 1024 * 1024 || total > 32 * 1024 * 1024) throw new InvalidDataException("Rule bundle exceeds its uncompressed size limit.");
            using Stream stream = entry.Open(); using var output = new MemoryStream(); stream.CopyTo(output);
            files.Add(entry.FullName, output.ToArray());
        }
        return new InvoiceRuleBundle(files, peppol);
    }

    /// <summary>True when the pinned optional Peppol rule source is loaded.</summary>
    public bool HasPeppolRules => _peppol != null;
    internal byte[] File(string path) => _files.TryGetValue(path, out byte[]? bytes) ? bytes : throw new InvalidDataException("Pinned bundle resource is missing: " + path);
    internal byte[] Peppol => _peppol ?? throw new InvalidOperationException("Load the pinned Peppol Schematron source to validate Peppol BIS.");
    internal string Schema(InvoiceSyntax syntax, bool credit) => syntax == InvoiceSyntax.Cii
        ? "resources/cii/16b/xsd/CrossIndustryInvoice_100pD16B.xsd"
        : "resources/ubl/2.1/xsd/maindoc/UBL-" + (credit ? "CreditNote" : "Invoice") + "-2.1.xsd";
    internal IEnumerable<(string Name, byte[] Bytes, bool Compile)> Rules(InvoiceRulesRelease release, InvoiceSyntax syntax) {
        string syntaxName = syntax == InvoiceSyntax.Cii ? "CII" : "UBL";
        string prefix = syntax == InvoiceSyntax.Cii ? "resources/cii/16b/xsl/" : "resources/ubl/2.1/xsl/";
        yield return ("EN16931-1.3.16-" + syntaxName, File(prefix + "EN16931-" + syntaxName + "-validation.xsl"), false);
        if (release == InvoiceRulesRelease.XRechnung_3_0_2_2026_08_31)
            yield return ("XRechnung-3.0.2-2026-08-31-" + syntaxName, File("resources/xrechnung/3.0.2/xsl/XRechnung-" + syntaxName + "-validation.xsl"), false);
        if (release == InvoiceRulesRelease.PeppolBis_3_0_21) yield return ("Peppol-BIS-3.0.21-UBL", Peppol, true);
    }
    internal IReadOnlyDictionary<string, InvoiceDiagnosticSeverity> SeverityOverrides(InvoiceRulesRelease release, InvoiceSyntax syntax, bool credit) {
        var result = new Dictionary<string, InvoiceDiagnosticSeverity>(StringComparer.Ordinal);
        if (release != InvoiceRulesRelease.XRechnung_3_0_2_2026_08_31) return result;
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
}
