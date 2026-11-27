using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    internal sealed partial class LegacyDocDocxWriter {
        private const string CorePropertiesPartName = "/docProps/core.xml";
        private const string CorePropertiesContentType = "application/vnd.openxmlformats-package.core-properties+xml";
        private const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private const string ApplicationPropertiesPartName = "/docProps/app.xml";
        private const string ApplicationPropertiesContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        private const string ApplicationPropertiesRelationshipType = RelationshipTypePrefix + "extended-properties";
        private const string XmlSchemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance";

        /// <summary>
        /// Writes the summary and document summary properties the source carries: the core properties part and the
        /// company and manager of the extended properties part. Custom properties are left to the caller.
        /// </summary>
        private void AddDocumentProperties() {
            LegacyDocDocumentProperties properties = _source.DocumentProperties;
            if (!properties.HasAnyProperties) return;
            AddCoreProperties(properties);
            if (properties.Company != null || properties.Manager != null) AddApplicationProperties(properties);
        }

        private void AddCoreProperties(LegacyDocDocumentProperties properties) {
            var values = new List<(int Namespace, string Name, string Value, bool Date)>();
            void Add(int namespaceId, string name, string? value) {
                if (value != null) values.Add((namespaceId, name, value, false));
            }

            void AddDate(int namespaceId, string name, DateTime? value, bool typed) {
                // The W3CDTF form the package writer uses: universal time with seven fraction digits.
                if (value != null) values.Add((namespaceId, name, value.Value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture), typed));
            }

            // The order the package writer stores them in.
            Add(OfficeXmlNamespaces.DublinCore, "title", properties.Title);
            Add(OfficeXmlNamespaces.DublinCore, "subject", properties.Subject);
            Add(OfficeXmlNamespaces.DublinCore, "creator", properties.Creator);
            Add(OfficeXmlNamespaces.CoreProperties, "keywords", properties.Keywords);
            Add(OfficeXmlNamespaces.DublinCore, "description", properties.Description);
            Add(OfficeXmlNamespaces.CoreProperties, "category", properties.Category);
            Add(OfficeXmlNamespaces.CoreProperties, "lastModifiedBy", properties.LastModifiedBy);
            Add(OfficeXmlNamespaces.CoreProperties, "revision", properties.Revision);
            AddDate(OfficeXmlNamespaces.DublinCoreTerms, "created", properties.Created, typed: true);
            AddDate(OfficeXmlNamespaces.DublinCoreTerms, "modified", properties.Modified, typed: true);
            AddDate(OfficeXmlNamespaces.CoreProperties, "lastPrinted", properties.LastPrinted, typed: false);
            if (values.Count == 0) return;

            OfficeXmlTree tree = _package.CreateXmlPart(CorePropertiesPartName, CorePropertiesContentType, OfficeXmlNamespaces.CoreProperties, "coreProperties");
            tree.DeclareNamespace(tree.Root, "cp", OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.CoreProperties)!);
            tree.DeclareNamespace(tree.Root, "dc", OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.DublinCore)!);
            tree.DeclareNamespace(tree.Root, "dcterms", OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.DublinCoreTerms)!);
            tree.DeclareNamespace(tree.Root, "xsi", XmlSchemaInstanceNamespace);
            int schemaInstance = tree.Namespaces.GetId(XmlSchemaInstanceNamespace);
            foreach ((int namespaceId, string name, string value, bool date) in values) {
                int element = tree.CreateElement(namespaceId, name);
                if (date) tree.SetAttribute(element, schemaInstance, "type", "dcterms:W3CDTF");
                tree.Append(element, tree.CreateText(value));
                tree.Append(tree.Root, element);
            }

            _package.AddRelationship(null, CorePropertiesRelationshipType, CorePropertiesPartName, id: NextRelationshipId());
        }

        private void AddApplicationProperties(LegacyDocDocumentProperties properties) {
            OfficeXmlTree tree = _package.CreateXmlPart(ApplicationPropertiesPartName, ApplicationPropertiesContentType, OfficeXmlNamespaces.ExtendedProperties, "Properties");
            tree.DeclareNamespace(tree.Root, string.Empty, OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.ExtendedProperties)!);
            void Add(string name, string? value) {
                if (value == null) return;
                int element = tree.CreateElement(OfficeXmlNamespaces.ExtendedProperties, name);
                if (value.Length > 0) tree.Append(element, tree.CreateText(value));
                tree.Append(tree.Root, element);
            }

            Add("Company", properties.Company);
            Add("Manager", properties.Manager);
            _package.AddRelationship(null, ApplicationPropertiesRelationshipType, ApplicationPropertiesPartName, id: NextRelationshipId());
        }
    }
}
