using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Internal.Invoicing;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Invoicing.KSeF;

/// <summary>Exact UPO bytes with separate schema, identity/hash binding and retrieval evidence. This does not verify a digital signature or fiscal treatment.</summary>
public sealed class KsefReceipt {
    private readonly byte[] _bytes;
    private KsefReceipt(byte[] bytes, bool authenticated, bool bound, InvoiceValidationStatus schemaStatus, IReadOnlyList<InvoiceDiagnostic> diagnostics) {
        _bytes = bytes; RetrievedThroughAuthenticatedApi = authenticated; IsBound = bound; SchemaStatus = schemaStatus; Diagnostics = diagnostics;
        Sha256 = Convert.ToHexString(SHA256.HashData(bytes));
    }
    /// <summary>Whether this instance was obtained through the client's direct authenticated official API route.</summary>
    public bool RetrievedThroughAuthenticatedApi { get; }
    /// <summary>Whether session, context, KSeF identifier and exact document hash match the expected values.</summary>
    public bool IsBound { get; }
    /// <summary>Pinned UPO v4-3 XSD result, independent of retrieval and invoice acceptance.</summary>
    public InvoiceValidationStatus SchemaStatus { get; }
    /// <summary>Schema and binding diagnostics; a failed schema is never silently relaxed.</summary>
    public IReadOnlyList<InvoiceDiagnostic> Diagnostics { get; }
    /// <summary>SHA-256 of exact UPO bytes, hexadecimal encoded.</summary>
    public string Sha256 { get; }
    /// <summary>Returns a defensive copy of the receipt as delivered.</summary>
    public byte[] GetBytes() => (byte[])_bytes.Clone();
    /// <summary>Inspects offline receipt bytes without claiming authenticated retrieval.</summary>
    public static KsefReceipt Inspect(byte[] xml, KsefContext context, string sessionReference, string ksefNumber, string invoiceHash, CancellationToken cancellationToken = default) =>
        Read(xml, context, sessionReference, ksefNumber, invoiceHash, false, cancellationToken);

    internal static KsefReceipt Read(byte[] xml, KsefContext context, string sessionReference, string ksefNumber, string invoiceHash, bool authenticated, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(xml); ArgumentNullException.ThrowIfNull(context); cancellationToken.ThrowIfCancellationRequested();
        KsefProtocol.Reference(sessionReference); KsefProtocol.InvoiceNumber(ksefNumber); KsefProtocol.Hash(invoiceHash);
        if (xml.Length == 0 || xml.Length > 2 * 1024 * 1024) throw new InvalidDataException("UPO must contain between one byte and 2 MiB.");
        byte[] bytes = (byte[])xml.Clone(); var diagnostics = new List<InvoiceDiagnostic>(); bool bound = false;
        try {
            XDocument document = InvoiceXml.Parse(bytes, InvoiceXml.MaximumNodes, InvoiceXml.MaximumAttributes);
            XNamespace ns = KsefProtocolSchemas.UpoNamespace;
            XElement root = document.Root ?? throw new InvalidDataException("UPO root is missing.");
            if (root.Name != ns + "Potwierdzenie") throw new InvalidDataException("Expected the pinned UPO v4-3 namespace.");
            XElement? authentication = InvoiceXml.Unique(root, ns + "Uwierzytelnienie");
            XElement? identifier = authentication == null ? null : InvoiceXml.Unique(authentication, ns + "IdKontekstu");
            string contextName = context.Kind switch { KsefContextKind.Nip => "Nip", KsefContextKind.InternalId => "IdWewnetrzny", KsefContextKind.NipVatUe => "IdZlozonyVatUE", KsefContextKind.PeppolId => "IdDostawcyUslugPeppol", _ => throw new InvalidDataException("Unsupported context kind.") };
            bool contextMatches = identifier?.Elements().Count() == 1 && Value(identifier, ns + contextName) == context.Value;
            XElement[] matches = root.Elements(ns + "Dokument").Where(element => Value(element, ns + "NumerKSeFDokumentu") == ksefNumber).Take(2).ToArray();
            bound = Value(root, ns + "NumerReferencyjnySesji") == sessionReference && contextMatches && matches.Length == 1 &&
                Value(matches[0], ns + "SkrotDokumentu") == invoiceHash && Value(root, ns + "KodFormularza") == "FA (3)";
            if (!bound) diagnostics.Add(new InvoiceDiagnostic("KSEF-UPO-BINDING", "Receipt does not uniquely match the expected session, context, KSeF number and invoice hash.", "Receipt"));
            IReadOnlyList<InvoiceDiagnostic> schema = KsefProtocolSchemas.ValidateUpo(bytes, cancellationToken); diagnostics.AddRange(schema);
            return new KsefReceipt(bytes, authenticated, bound, schema.Any(item => item.Severity == InvoiceDiagnosticSeverity.Error) ? InvoiceValidationStatus.Invalid : InvoiceValidationStatus.Passed, diagnostics.AsReadOnly());
        } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
            diagnostics.Add(new InvoiceDiagnostic("KSEF-UPO-XML", "Receipt XML is malformed, ambiguous or outside the supported bounds.", "Receipt"));
            return new KsefReceipt(bytes, authenticated, false, InvoiceValidationStatus.Invalid, diagnostics.AsReadOnly());
        }
    }
    private static string? Value(XElement parent, XName name) {
        XElement? element = InvoiceXml.Unique(parent, name); return element == null ? null : InvoiceXml.Scalar(element);
    }
}
