using System.Net;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Validation;

internal static class InvoiceSchemaValidation {
    internal static List<InvoiceDiagnostic> Validate(byte[] xml, InvoiceRuleBundle bundle, InvoiceSyntax syntax, bool credit, CancellationToken cancellationToken) {
        var resolver = new BundleResolver(bundle);
        var schemas = new XmlSchemaSet { XmlResolver = resolver };
        string schemaPath = bundle.Schema(syntax, credit);
        using var schemaInput = new MemoryStream(bundle.File(schemaPath), false);
        using (XmlReader schemaReader = XmlReader.Create(schemaInput, InvoiceRuleBundle.XmlSettings(), "invoice-bundle:///" + schemaPath)) schemas.Add(null, schemaReader);
        schemas.Compile();
        var diagnostics = new InvoiceDiagnosticBuffer();
        XmlReaderSettings settings = InvoiceRuleBundle.XmlSettings();
        settings.Schemas = schemas; settings.ValidationType = ValidationType.Schema;
        settings.ValidationFlags = XmlSchemaValidationFlags.ReportValidationWarnings | XmlSchemaValidationFlags.ProcessIdentityConstraints;
        settings.ValidationEventHandler += (_, args) => {
            diagnostics.Add(new InvoiceDiagnostic("XSD", args.Message, "line " + args.Exception.LineNumber + ":" + args.Exception.LinePosition,
                args.Severity == XmlSeverityType.Warning ? InvoiceDiagnosticSeverity.Warning : InvoiceDiagnosticSeverity.Error));
        };
        using var input = new MemoryStream(xml, false);
        using XmlReader reader = XmlReader.Create(input, settings);
        while (reader.Read()) cancellationToken.ThrowIfCancellationRequested();
        return diagnostics.ToList();
    }
    private sealed class BundleResolver(InvoiceRuleBundle bundle) : XmlResolver {
        public override ICredentials? Credentials { set { } }
        public override object GetEntity(Uri absoluteUri, string? role, Type? ofObjectToReturn) {
            if (absoluteUri.Scheme != "invoice-bundle" || absoluteUri.Host.Length != 0 || absoluteUri.Query.Length != 0 || absoluteUri.Fragment.Length != 0)
                throw new XmlException("Schema resolution is restricted to the pinned local bundle.");
            byte[] schema = bundle.File(Uri.UnescapeDataString(absoluteUri.AbsolutePath.TrimStart('/')));
            // The pinned OASIS XMLDSig schema contains an internal DTD. Expand its bounded
            // internal declarations only; invoice XML still rejects every DTD.
            using var input = new MemoryStream(schema, false);
            using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Parse, XmlResolver = null, MaxCharactersInDocument = 4 * 1024 * 1024,
                MaxCharactersFromEntities = 65536
            });
            XDocument document = XDocument.Load(reader);
            document.DocumentType?.Remove();
            var output = new MemoryStream(); document.Save(output); output.Position = 0;
            return output;
        }
    }
}
