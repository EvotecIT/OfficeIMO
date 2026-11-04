using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private static readonly XNamespace CoreNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private static readonly XNamespace DublinCoreNamespace = "http://purl.org/dc/elements/1.1/";
        private static readonly XNamespace DublinTermsNamespace = "http://purl.org/dc/terms/";
        private static readonly XNamespace SchemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance";

        /// <summary>Copies a loaded package and replaces only its core properties when nothing else changed.</summary>
        private bool TryWritePropertiesOnlyPackage(Stream destination, ExcelSaveOptions? options) {
            byte[]? source = _propertiesOnlySourceBytes;
            if (source == null
                || options?.DisableFastPackageWriter == true
                || options?.SafePreflight == true
                || options?.SafeRepairDefinedNames == true
                || options?.ValidateOpenXml == true
                || HasCalculationSaveWork(options)
                || _spreadSheetDocument.DigitalSignatureOriginPart != null
                || _spreadSheetDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null) {
                return false;
            }

            _spreadSheetDocument.PackageProperties.Modified = DateTime.UtcNow;
            byte[]? package = ExcelPackageEntrySplicer.TryEdit(source, splicer => {
                KeyValuePair<string, byte[]> core = RewriteCoreProperties(splicer);
                splicer.Replace(core.Key, core.Value);
            });
            if (package == null) return false;

            ThrowIfPackageMaterializationExceedsLimit(package.LongLength, options);
            PrepareDestinationStreamForWrite(destination);
            destination.Write(package, 0, package.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            MarkPackageClean(package);
            return true;
        }

        private KeyValuePair<string, byte[]> RewriteCoreProperties(ExcelPackageEntrySplicer splicer) {
            string entryName = FindCorePropertiesEntry(splicer.Read("_rels/.rels"));
            XDocument document = LoadPackageXml(splicer.Read(entryName));
            XElement root = document.Root ?? throw new InvalidDataException("Core properties are empty.");
            var properties = _spreadSheetDocument.PackageProperties;
            SetCoreText(root, DublinCoreNamespace + "title", properties.Title);
            SetCoreText(root, DublinCoreNamespace + "creator", properties.Creator);
            SetCoreText(root, DublinCoreNamespace + "subject", properties.Subject);
            SetCoreText(root, CoreNamespace + "category", properties.Category);
            SetCoreText(root, DublinCoreNamespace + "description", properties.Description);
            SetCoreText(root, CoreNamespace + "keywords", properties.Keywords);
            SetCoreText(root, CoreNamespace + "lastModifiedBy", properties.LastModifiedBy);
            SetCoreText(root, CoreNamespace + "version", properties.Version);
            SetCoreDate(root, DublinTermsNamespace + "created", properties.Created, typed: true);
            SetCoreDate(root, DublinTermsNamespace + "modified", properties.Modified, typed: true);
            SetCoreDate(root, CoreNamespace + "lastPrinted", properties.LastPrinted, typed: false);

            using var output = new MemoryStream();
            using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false) })) {
                document.Save(writer);
            }
            return new KeyValuePair<string, byte[]>(entryName, output.ToArray());
        }

        private static XDocument LoadPackageXml(byte[] xml) {
            using var input = new MemoryStream(xml, writable: false);
            using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
            return XDocument.Load(reader);
        }

        private static string FindCorePropertiesEntry(byte[] relationships) {
            XNamespace package = "http://schemas.openxmlformats.org/package/2006/relationships";
            XElement relationship = LoadPackageXml(relationships).Root?.Elements(package + "Relationship")
                .SingleOrDefault(candidate => string.Equals((string?)candidate.Attribute("Type"), CorePropertiesRelationshipType, StringComparison.Ordinal))
                ?? throw new InvalidDataException("The package has no core properties.");
            return ((string?)relationship.Attribute("Target"))?.TrimStart('/')
                ?? throw new InvalidDataException("The core properties relationship has no target.");
        }

        private static void SetCoreText(XElement root, XName name, string? value) {
            XElement? element = root.Element(name);
            if (value == null) {
                element?.Remove();
            } else if (element == null) {
                root.Add(new XElement(name, value));
            } else {
                element.Value = value;
            }
        }

        private static void SetCoreDate(XElement root, XName name, DateTime? value, bool typed) {
            if (value == null) {
                root.Element(name)?.Remove();
                return;
            }

            DateTime utc = value.Value.Kind == DateTimeKind.Local ? value.Value.ToUniversalTime() : value.Value;
            SetCoreText(root, name, utc.ToString("yyyy-MM-ddTHH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture));
            if (typed) {
                root.Element(name)!.SetAttributeValue(SchemaInstanceNamespace + "type", "dcterms:W3CDTF");
            }
        }
    }
}
