using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using OfficeIMO.Internal.Invoicing;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Invoicing.KSeF;

internal static class KsefProtocolSchemas {
    internal const string AuthNamespace = "http://ksef.mf.gov.pl/auth/token/2.1";
    internal const string UpoNamespace = "http://upo.schematy.mf.gov.pl/KSeF/v4-3";
    private static readonly Lazy<XmlSchemaSet> Auth = new(() => Load("auth-v2-1.xsd", "D48477808E342892F62EF8BC7F0B134812F0536B3BB9CC8433D4FFF2A926C44C"));
    private static readonly Lazy<XmlSchemaSet> Upo = new(() => Load("upo-v4-3.xsd", "1E5FF386A29324021A9E0126319680AEC0B1E0D4F4A18ADD30B2F5D12CE6FA86"));

    internal static void CheckContext(KsefContext context) {
        XNamespace ns = AuthNamespace;
        var xml = new XDocument(new XElement(ns + "AuthTokenRequest", new XElement(ns + "Challenge", "20260930-CR-0000000000-0000000000-00"),
            new XElement(ns + "ContextIdentifier", new XElement(ns + context.Kind.ToString(), context.Value)), new XElement(ns + "SubjectIdentifierType", "certificateSubject")));
        List<InvoiceDiagnostic> diagnostics = InvoiceSchemaValidation.ValidateDocument(Encoding.UTF8.GetBytes(xml.ToString()), Auth.Value, default);
        if (diagnostics.Any(item => item.Severity == InvoiceDiagnosticSeverity.Error)) throw new ArgumentException("Context identifier does not satisfy the pinned authentication schema.", nameof(context));
    }
    internal static IReadOnlyList<InvoiceDiagnostic> ValidateUpo(byte[] bytes, CancellationToken cancellationToken) => InvoiceSchemaValidation.ValidateDocument(bytes, Upo.Value, cancellationToken).AsReadOnly();
    private static XmlSchemaSet Load(string name, string hash) {
        using Stream input = typeof(KsefProtocolSchemas).Assembly.GetManifestResourceStream("OfficeIMO.Invoicing.KSeF.Schemas." + name) ?? throw new InvalidDataException("Pinned protocol schema is missing.");
        using var output = new MemoryStream(); input.CopyTo(output); byte[] bytes = output.ToArray();
        if (bytes.Length > 65_536 || Convert.ToHexString(SHA256.HashData(bytes)) != hash) throw new InvalidDataException("Pinned protocol schema failed its integrity check.");
        using var buffer = new MemoryStream(bytes, false);
        using XmlReader reader = XmlReader.Create(buffer, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 65_536 });
        var schemas = new XmlSchemaSet { XmlResolver = null }; schemas.Add(null, reader); schemas.Compile(); return schemas;
    }
}
