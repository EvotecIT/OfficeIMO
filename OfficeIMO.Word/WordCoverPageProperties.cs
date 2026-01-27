using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;

namespace OfficeIMO.Word {
    /// <summary>
    /// Manages the cover page properties custom XML part used by Word's built-in
    /// cover page templates (store item ID: {55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}).
    /// </summary>
    public class WordCoverPageProperties {
        private const string CoverPagePropsNamespace = "http://schemas.microsoft.com/office/2006/coverPageProps";
        internal const string CoverPagePropsStoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}";

        private static readonly XNamespace CoverPagePropsNs = CoverPagePropsNamespace;
        private readonly WordDocument _document;

        /// <summary>
        /// Initializes a new instance for the provided document.
        /// </summary>
        /// <param name="document">Document that owns the cover page properties.</param>
        public WordCoverPageProperties(WordDocument document) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// Gets or sets the publish date used by cover page templates that bind to PublishDate.
        /// </summary>
        public string PublishDate {
            get => Get("PublishDate") ?? string.Empty;
            set => Set("PublishDate", value);
        }

        /// <summary>
        /// Gets or sets the abstract used by cover page templates that bind to Abstract.
        /// </summary>
        public string Abstract {
            get => Get("Abstract") ?? string.Empty;
            set => Set("Abstract", value);
        }

        /// <summary>
        /// Gets or sets the company address used by cover page templates that bind to CompanyAddress.
        /// </summary>
        public string CompanyAddress {
            get => Get("CompanyAddress") ?? string.Empty;
            set => Set("CompanyAddress", value);
        }

        /// <summary>
        /// Gets or sets the company email used by cover page templates that bind to CompanyEmail.
        /// </summary>
        public string CompanyEmail {
            get => Get("CompanyEmail") ?? string.Empty;
            set => Set("CompanyEmail", value);
        }

        /// <summary>
        /// Retrieves a cover page property value by element name.
        /// </summary>
        /// <param name="name">Element name within the coverPageProps namespace.</param>
        /// <returns>The stored value when present; otherwise <c>null</c>.</returns>
        public string? Get(string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                return null;
            }

            var part = FindCoverPagePropsPart();
            if (part == null) {
                return null;
            }

            var document = LoadXml(part);
            var root = EnsureRoot(document);
            return root.Element(CoverPagePropsNs + name)?.Value;
        }

        /// <summary>
        /// Sets a cover page property value by element name, creating the custom XML part when required.
        /// </summary>
        /// <param name="name">Element name within the coverPageProps namespace.</param>
        /// <param name="value">Value to store.</param>
        public void Set(string name, string? value) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Property name cannot be null or empty.", nameof(name));
            }

            var part = EnsureCoverPagePropsPart();
            var document = LoadXml(part);
            var root = EnsureRoot(document);
            var element = root.Element(CoverPagePropsNs + name);
            if (element == null) {
                element = new XElement(CoverPagePropsNs + name);
                root.Add(element);
            }

            element.Value = value ?? string.Empty;
            SaveXml(part, document);

            // Encourage Word to refresh bound content controls on open.
            _document.Settings.UpdateFieldsOnOpen = true;
        }

        private CustomXmlPart? FindCoverPagePropsPart() {
            var mainPart = _document._wordprocessingDocument?.MainDocumentPart;
            if (mainPart == null) {
                return null;
            }

            foreach (var customXmlPart in mainPart.CustomXmlParts) {
                var itemId = customXmlPart.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value;
                if (string.Equals(itemId, CoverPagePropsStoreItemId, StringComparison.OrdinalIgnoreCase)) {
                    return customXmlPart;
                }
            }

            return null;
        }

        private CustomXmlPart EnsureCoverPagePropsPart() {
            var existing = FindCoverPagePropsPart();
            if (existing != null) {
                EnsureCoverPagePropsSchemaReference(existing);
                EnsureCoverPagePropsXml(existing);
                return existing;
            }

            var mainPart = _document._wordprocessingDocument?.MainDocumentPart
                ?? throw new InvalidOperationException("Main document part is missing.");

            var part = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            var propertiesPart = part.AddNewPart<CustomXmlPropertiesPart>();

            var dataStoreItem = new DataStoreItem { ItemId = CoverPagePropsStoreItemId };
            var schemaReferences = new SchemaReferences();
            schemaReferences.Append(new SchemaReference { Uri = CoverPagePropsNamespace });
            dataStoreItem.Append(schemaReferences);
            propertiesPart.DataStoreItem = dataStoreItem;

            EnsureCoverPagePropsXml(part);
            return part;
        }

        private static void EnsureCoverPagePropsSchemaReference(CustomXmlPart part) {
            var dataStoreItem = part.CustomXmlPropertiesPart?.DataStoreItem;
            if (dataStoreItem == null) {
                return;
            }

            var schemaReferences = dataStoreItem.GetFirstChild<SchemaReferences>();
            if (schemaReferences == null) {
                schemaReferences = new SchemaReferences();
                dataStoreItem.Append(schemaReferences);
            }

            var hasCoverPageSchema = schemaReferences
                .Elements<SchemaReference>()
                .Any(r => string.Equals(r.Uri?.Value, CoverPagePropsNamespace, StringComparison.OrdinalIgnoreCase));

            if (!hasCoverPageSchema) {
                schemaReferences.Append(new SchemaReference { Uri = CoverPagePropsNamespace });
            }
        }

        private static void EnsureCoverPagePropsXml(CustomXmlPart part) {
            var hasContent = false;
            using (var readStream = part.GetStream(FileMode.OpenOrCreate, FileAccess.Read)) {
                hasContent = readStream.Length > 0;
            }

            if (hasContent) {
                return;
            }

            var document = new XDocument(new XElement(CoverPagePropsNs + "CoverPageProperties"));
            using var writeStream = part.GetStream(FileMode.Create, FileAccess.Write);
            document.Save(writeStream);
        }

        private static XDocument LoadXml(CustomXmlPart part) {
            using var stream = part.GetStream(FileMode.OpenOrCreate, FileAccess.Read);
            if (stream.Length == 0) {
                return new XDocument(new XElement(CoverPagePropsNs + "CoverPageProperties"));
            }

            try {
                return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            } catch {
                // If the part is malformed, fall back to a minimal valid structure.
                return new XDocument(new XElement(CoverPagePropsNs + "CoverPageProperties"));
            }
        }

        private static void SaveXml(CustomXmlPart part, XDocument document) {
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            document.Save(stream);
        }

        private static XElement EnsureRoot(XDocument document) {
            if (document.Root == null) {
                document.Add(new XElement(CoverPagePropsNs + "CoverPageProperties"));
            }

            if (document.Root!.Name != CoverPagePropsNs + "CoverPageProperties") {
                var newRoot = new XElement(CoverPagePropsNs + "CoverPageProperties");
                newRoot.Add(document.Root.Nodes());
                document.Root.ReplaceWith(newRoot);
            }

            return document.Root!;
        }
    }
}
