using OfficeIMO.Invoicing;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Pdf;

/// <summary>Connects the invoice owner's profile declaration to PDF attachment/XMP invariants.</summary>
internal static class PdfElectronicInvoiceProfile {
    /// <summary>Checks original RDF scalar values before convenience readback values are normalized.</summary>
    internal static string? GetRawXmpDiagnostic(string? rawXml) {
        if (string.IsNullOrEmpty(rawXml)) return "The original invoice XMP packet is unavailable.";
        try {
            using var input = new StringReader(rawXml!);
            using var reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16 * 1024 * 1024
            });
            XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            XNamespace rdf = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
            XNamespace fx = PdfElectronicInvoiceMetadata.FacturXNamespaceUri;
            XElement[] descriptions = document.Descendants(rdf + "Description")
                .Where(e => e.Parent?.Name == rdf + "RDF" && string.IsNullOrEmpty((string?)e.Attribute(rdf + "about"))).ToArray();
            foreach (string name in new[] { "DocumentType", "DocumentFileName", "Version", "ConformanceLevel" }) {
                var values = new List<string>();
                foreach (XElement description in descriptions) {
                    foreach (XElement property in description.Elements(fx + name)) {
                        if (property.HasElements || property.Attributes().Any(a => !a.IsNamespaceDeclaration))
                            return "Invoice XMP " + name + " must be an unqualified scalar property.";
                        values.Add(property.Value);
                    }
                    values.AddRange(description.Attributes(fx + name).Select(a => a.Value));
                }
                if (values.Count != 1 || string.IsNullOrWhiteSpace(values[0]) || values[0] != values[0].Trim())
                    return "Invoice XMP " + name + " requires one canonical scalar declaration without surrounding whitespace.";
            }
            return null;
        } catch (XmlException exception) {
            return "Invoice XMP is malformed: " + exception.Message;
        }
    }

    internal static PdfElectronicInvoiceMetadata CreateMetadata(PdfEmbeddedFile attachment, string? requestedLevel, string version) {
        InvoiceProfileDeclaration declaration = InvoiceProfileDeclaration.Read(attachment.DataSnapshot);
        if (declaration.Syntax != InvoiceSyntax.Cii || !declaration.Profile.HasValue || declaration.Profile == InvoiceProfile.PeppolBis)
            throw new ArgumentException("Factur-X requires a supported CII guideline identifier. Found: " + declaration.GuidelineId + ".", nameof(attachment));
        InvoiceProfile profile = declaration.Profile.Value;
        string canonical = InvoiceProfiles.GetXmpConformanceLevel(profile);
        if (requestedLevel != null && (!InvoiceProfiles.TryFromXmpConformanceLevel(requestedLevel, out InvoiceProfile requested) || requested != profile))
            throw new ArgumentException("Factur-X XMP profile '" + requestedLevel + "' does not match XML guideline '" + declaration.GuidelineId + "' (" + canonical + ").", nameof(requestedLevel));
        if (!string.Equals(version, "1.0", StringComparison.Ordinal))
            throw new ArgumentException("Factur-X XMP Version is 1.0; it is not the invoice specification release number.", nameof(version));
        return PdfElectronicInvoiceMetadata.FacturX(canonical, version);
    }

    /// <summary>Returns a diagnostic for any missing, ambiguous, or conflicting declaration.</summary>
    internal static string? GetDiagnostic(PdfElectronicInvoiceMetadata? metadata, IReadOnlyList<PdfEmbeddedFile> attachments) {
        if (metadata == null) return "Factur-X requires XMP invoice metadata.";
        if (metadata.DocumentType != "INVOICE" || metadata.DocumentFileName != "factur-x.xml" || metadata.Version != "1.0")
            return "Factur-X XMP must declare DocumentType=INVOICE, DocumentFileName=factur-x.xml, and Version=1.0.";
        PdfEmbeddedFile[] invoices = attachments.Where(file => file.FileName == "factur-x.xml").ToArray();
        if (invoices.Length != 1) return "Factur-X requires exactly one canonical factur-x.xml attachment; found " + invoices.Length + ".";
        try {
            PdfElectronicInvoiceMetadata expected = CreateMetadata(invoices[0], metadata.ConformanceLevel, metadata.Version);
            if (metadata.ConformanceLevel != expected.ConformanceLevel)
                return "Factur-X XMP ConformanceLevel must use the canonical spelling " + expected.ConformanceLevel + ".";
            return null;
        } catch (Exception exception) when (exception is ArgumentException || exception is InvalidDataException || exception is System.Xml.XmlException) {
            return exception.Message;
        }
    }
}
