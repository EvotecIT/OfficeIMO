using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Security.Cryptography;
using OfficeIMO.Internal.Invoicing;

namespace OfficeIMO.Invoicing.Validation;

/// <summary>Schema-only qualification of exact FA(3) bytes. This result does not certify fiscal correctness or KSeF acceptance.</summary>
public sealed class Fa3SchemaValidationResult {
    internal Fa3SchemaValidationResult(InvoiceValidationStatus status, IReadOnlyList<InvoiceDiagnostic> diagnostics, int length, string? hash) {
        Status = status; Diagnostics = diagnostics; InputByteCount = length; InputSha256 = hash;
    }
    /// <summary>Whether the bounded XSD stage passed, rejected input or could not complete.</summary>
    public InvoiceValidationStatus Status { get; }
    /// <summary>True only when the exact document passed the pinned schema.</summary>
    public bool IsValid => Status == InvoiceValidationStatus.Passed;
    /// <summary>Identity of the main schema used by this result.</summary>
    public string SchemaSha256 => Fa3SchemaBundle.SchemaSha256;
    /// <summary>Length of the observed input bytes.</summary>
    public int InputByteCount { get; }
    /// <summary>SHA-256 of the defensive input snapshot, or null when its byte length violates the input bound.</summary>
    public string? InputSha256 { get; }
    /// <summary>Bounded errors and warnings. Truncation never changes failure to success.</summary>
    public IReadOnlyList<InvoiceDiagnostic> Diagnostics { get; }
}

/// <summary>Offline, bounded FA(3) XSD validation over an immutable pinned schema bundle.</summary>
public sealed class Fa3SchemaValidator {
    private readonly Fa3SchemaBundle _bundle;
    /// <summary>Creates a validator using a verified bundle; no schema download occurs.</summary>
    public Fa3SchemaValidator(Fa3SchemaBundle bundle) => _bundle = bundle ?? throw new ArgumentNullException(nameof(bundle));

    /// <summary>Validates a defensive input snapshot, rejecting DTDs and document-controlled schema locations.</summary>
    public Fa3SchemaValidationResult Validate(byte[] xml, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(xml); cancellationToken.ThrowIfCancellationRequested();
        if (xml.Length == 0 || xml.Length > InvoiceProfileDeclaration.MaximumXmlBytes)
            return Invalid("FA3-INPUT", "Invoice XML must contain between 1 byte and 16 MiB.", xml.Length, null);
        byte[] snapshot = (byte[])xml.Clone();
        string hash = Convert.ToHexString(SHA256.HashData(snapshot));
        try {
            XDocument document = InvoiceXml.Parse(snapshot);
            if (document.Root?.Name != XName.Get("Faktura", Fa3InvoiceReader.NamespaceUri))
                return Invalid("FA3-ROOT", "Expected the pinned FA(3) Faktura namespace.", snapshot.Length, hash);
            var diagnostics = InvoiceSchemaValidation.ValidateDocument(snapshot, _bundle.CreateSchemas(), cancellationToken);
            return new Fa3SchemaValidationResult(diagnostics.Any(item => item.Severity == InvoiceDiagnosticSeverity.Error)
                ? InvoiceValidationStatus.Invalid : InvoiceValidationStatus.Passed, diagnostics.AsReadOnly(), snapshot.Length, hash);
        } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
            return Invalid("FA3-INPUT", exception.Message, snapshot.Length, hash);
        } catch (XmlSchemaException exception) {
            return new Fa3SchemaValidationResult(InvoiceValidationStatus.Failed,
                new[] { new InvoiceDiagnostic("FA3-SCHEMA-ENGINE", exception.Message, "Schema") }, snapshot.Length, hash);
        }
    }
    private static Fa3SchemaValidationResult Invalid(string code, string message, int length, string? hash) =>
        new(InvoiceValidationStatus.Invalid, new[] { new InvoiceDiagnostic(code, message, "Source") }, length, hash);
}
