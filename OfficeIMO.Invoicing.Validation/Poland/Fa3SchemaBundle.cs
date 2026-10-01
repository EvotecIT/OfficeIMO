using System.Net;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Schema;

namespace OfficeIMO.Invoicing.Validation;

/// <summary>Immutable snapshot of the official FA(3) 1-0E schema and its three pinned imported schemas.</summary>
/// <remarks>Loading and validation never fetch resources from the network. Every input file must match the published pinned digest.</remarks>
public sealed class Fa3SchemaBundle {
    /// <summary>SHA-256 of the official main schema.</summary>
    public const string SchemaSha256 = "B646B6B525F51ADF1BB2545F111FC8CA6E7AA6DD2F98948F1667D3695C06D958";
    private const string DefinitionsBase = "http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2022/01/05/eD/DefinicjeTypy/";
    private static readonly (string Name, string Hash)[] Files = {
        ("FA3.xsd", SchemaSha256),
        ("StrukturyDanych_v10-0E.xsd", "1137CE6E3C11C2B9EF3F05E4E72D6DCD6B4FA94908EA558F2BA15DE0259BB2AA"),
        ("ElementarneTypyDanych_v10-0E.xsd", "8A531CB181D3E298D11B28766655AE91FEE2D7851440095932FFC82137ED2BE1"),
        ("KodyKrajow_v10-0E.xsd", "1D41A1B3184188F2D20A51D3AFDE26204DDA182EC5DACF018204DCC9870DC644")
    };
    private readonly Dictionary<string, byte[]> _files;
    private Fa3SchemaBundle(Dictionary<string, byte[]> files) => _files = files;

    /// <summary>Snapshots the four named XSD files from a local directory and verifies their exact bytes before use.</summary>
    public static Fa3SchemaBundle LoadDirectory(string directory) {
        ArgumentException.ThrowIfNullOrWhiteSpace(directory);
        var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach ((string name, string hash) in Files) {
            using var input = new FileStream(Path.Combine(directory, name), FileMode.Open, FileAccess.Read, FileShare.Read);
            if (input.Length is <= 0 or > 4 * 1024 * 1024) throw new InvalidDataException("FA(3) schema file must contain 1 byte to 4 MiB: " + name);
            byte[] bytes = new byte[input.Length]; input.ReadExactly(bytes);
            if (input.ReadByte() != -1 || Convert.ToHexString(SHA256.HashData(bytes)) != hash)
                throw new InvalidDataException("FA(3) schema SHA-256 does not match the pinned release: " + name);
            files.Add(name, bytes);
        }
        return new Fa3SchemaBundle(files);
    }

    internal XmlSchemaSet CreateSchemas() {
        var resolver = new PinnedResolver(_files);
        var schemas = new XmlSchemaSet { XmlResolver = resolver };
        using var input = new MemoryStream(_files["FA3.xsd"], false);
        using var reader = XmlReader.Create(input, InvoiceRuleBundle.XmlSettings(), "fa3-bundle:///FA3.xsd");
        schemas.Add(Fa3InvoiceReader.NamespaceUri, reader);
        schemas.Compile(); return schemas;
    }

    private sealed class PinnedResolver(Dictionary<string, byte[]> files) : XmlResolver {
        public override ICredentials? Credentials { set { } }
        public override object GetEntity(Uri absoluteUri, string? role, Type? ofObjectToReturn) {
            // Original absolute CRD import URIs resolve only to their verified local snapshots.
            string? name = Files.Select(file => file.Name).FirstOrDefault(file =>
                absoluteUri.AbsoluteUri == DefinitionsBase + file || absoluteUri.AbsoluteUri == "fa3-bundle:///" + file);
            if (name == null) throw new XmlException("FA(3) schema resolution is restricted to the four pinned local resources.");
            return new MemoryStream(files[name], false);
        }
    }
}
