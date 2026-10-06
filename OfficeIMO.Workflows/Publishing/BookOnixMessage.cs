using OfficeIMO.Provenance;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace OfficeIMO.Workflows;

/// <summary>A schema-validated ONIX 3.1 message composed from complete exported product records.</summary>
public sealed class BookOnixMessage {
    private BookOnixMessage(byte[] bytes, BookOnixExportResult[] products) {
        Bytes = bytes;
        Products = Array.AsReadOnly(products);
    }

    /// <summary>UTF-8 ONIX XML, at most 16 MiB, with products in the supplied order.</summary>
    public byte[] Bytes { get; }

    /// <summary>
    /// Original export results, retaining each exact EPUB, hash, writer report and import diagnostics.
    /// The collection is snapshotted, but the results' byte arrays remain caller-owned and mutable.
    /// Integrity is checked during composition; later edits do not update this message.
    /// </summary>
    public IReadOnlyList<BookOnixExportResult> Products { get; }

    /// <summary>
    /// Composes 1–1000 unmodified ExportOnix results with the same serialized header (sender and UTC timestamp).
    /// Rejects repeated record references or ISBNs, altered ONIX or EPUB bytes, and combined XML over 16 MiB.
    /// The supplied compiled schema validates the complete message. No input is changed and no delivery occurs.
    /// Callers own schema provenance and must not mutate inputs or schemas during this operation.
    /// </summary>
    public static BookOnixMessage Create(IReadOnlyList<BookOnixExportResult> products, XmlSchemaSet schemas,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(products);
        ArgumentNullException.ThrowIfNull(schemas);
        cancellationToken.ThrowIfCancellationRequested();
        RequireSchema(schemas);
        if (products.Count is < 1 or > 1000)
            throw new ArgumentException("Supply 1–1000 exported products.", nameof(products));
        var snapshot = products.ToArray();
        XNamespace ns = BookProject.OnixNamespace;
        var root = new XElement(ns + "ONIXMessage", new XAttribute("release", "3.1"));
        XElement? header = null;
        var references = new HashSet<string>(StringComparer.Ordinal);
        var isbns = new HashSet<string>(StringComparer.Ordinal);
        long inputBytes = 0;
        foreach (BookOnixExportResult product in snapshot) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(product);
            inputBytes += product.Bytes.LongLength;
            if (inputBytes > 16L * 1024 * 1024)
                throw new InvalidDataException("Combined source ONIX XML exceeds 16 MiB.");
            if (Convert.ToHexString(SHA256.HashData(product.Bytes)) != product.OnixSha256 ||
                Convert.ToHexString(SHA256.HashData(product.Publication.Bytes)) != product.PublicationSha256)
                throw new InvalidDataException("An exported ONIX record or its associated EPUB has been altered.");
            cancellationToken.ThrowIfCancellationRequested();
            using var input = new MemoryStream(product.Bytes, false);
            using var reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1024L * 1024
            });
            var document = XDocument.Load(reader);
            XElement sourceHeader = document.Root!.Element(ns + "Header")!;
            if (header == null) { header = sourceHeader; root.Add(new XElement(header)); }
            else if (!XNode.DeepEquals(header, sourceHeader))
                throw new ArgumentException("All products must have the same sender and message timestamp.", nameof(products));
            XElement record = document.Root.Element(ns + "Product")!;
            string reference = record.Element(ns + "RecordReference")!.Value;
            string isbn = record.Elements(ns + "ProductIdentifier").Single(e => e.Element(ns + "ProductIDType")!.Value == "15")
                .Element(ns + "IDValue")!.Value;
            if (!references.Add(reference) || !isbns.Add(isbn))
                throw new ArgumentException("Each product must have a distinct record reference and ISBN.", nameof(products));
            root.Add(new XElement(record));
        }
        byte[] bytes = ValidateAndSerialize(new XDocument(root), schemas, 16L * 1024 * 1024, cancellationToken);
        return new BookOnixMessage(bytes, snapshot);
    }

    internal static void RequireSchema(XmlSchemaSet schemas) {
        if (!schemas.IsCompiled || schemas.GlobalElements[new XmlQualifiedName("ONIXMessage", BookProject.OnixNamespace)] == null)
            throw new ArgumentException("Supply a compiled schema set declaring ONIX 3.1 reference ONIXMessage.", nameof(schemas));
    }

    internal static byte[] ValidateAndSerialize(XDocument message, XmlSchemaSet schemas, long maximumBytes,
        CancellationToken cancellationToken) {
        message.Validate(schemas, (_, args) => {
            cancellationToken.ThrowIfCancellationRequested();
            throw new InvalidDataException("ONIX schema validation failed: " + args.Message, args.Exception);
        });
        cancellationToken.ThrowIfCancellationRequested();
        using var output = new OfficeProvenanceBoundedMemoryStream(maximumBytes);
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false, true), CloseOutput = false }))
            message.Save(writer);
        cancellationToken.ThrowIfCancellationRequested();
        return output.ToArray();
    }
}
