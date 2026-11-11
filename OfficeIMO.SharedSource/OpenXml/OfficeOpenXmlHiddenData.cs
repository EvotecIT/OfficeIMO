using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using Ext = DocumentFormat.OpenXml.ExtendedProperties;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>Package-level hidden data shared by Word, Excel and PowerPoint: properties, thumbnail, custom XML, image metadata and printer settings.</summary>
    internal static class OfficeOpenXmlHiddenData {
        internal const string CorePropertiesLocation = "/docProps/core.xml";
        internal const string ExtendedPropertiesLocation = "/docProps/app.xml";
        internal const string CustomPropertiesLocation = "/docProps/custom.xml";

        /// <summary>Identity and tracking properties from the core and extended property parts.</summary>
        internal static void InspectDocumentProperties(OpenXmlPackage package, OfficeHiddenDataCollector collector) {
            const OfficeHiddenDataKind kind = OfficeHiddenDataKind.DocumentProperties;
            var core = package.PackageProperties;
            if (!string.IsNullOrEmpty(core.Creator)) collector.Add(kind, CorePropertiesLocation, "Author", core.Creator, () => core.Creator = null);
            if (!string.IsNullOrEmpty(core.LastModifiedBy)) collector.Add(kind, CorePropertiesLocation, "Last modified by", core.LastModifiedBy, () => core.LastModifiedBy = null);
            if (core.LastPrinted.HasValue) collector.Add(kind, CorePropertiesLocation, "Last printed", core.LastPrinted.Value.ToString("u", CultureInfo.InvariantCulture), () => core.LastPrinted = null);
            if (!string.IsNullOrEmpty(core.Revision)) collector.Add(kind, CorePropertiesLocation, "Revision number", core.Revision, () => core.Revision = null);

            Ext.Properties? extended = GetExtendedProperties(package);
            if (extended == null) return;
            AddExtended(extended.Company, "Company");
            AddExtended(extended.Manager, "Manager");
            AddExtended(extended.Template, "Template");
            if (extended.TotalTime?.Text != "0") AddExtended(extended.TotalTime, "Total editing time (minutes)");
            AddExtended(extended.HyperlinkBase, "Hyperlink base");

            void AddExtended(OpenXmlLeafTextElement? element, string description) {
                if (element == null || string.IsNullOrEmpty(element.Text)) return;
                collector.Add(kind, ExtendedPropertiesLocation, description, element.Text, element.Remove);
            }
        }

        /// <summary>Custom document properties. <paramref name="removeFromModel"/> keeps a format's in-memory property collection in step.</summary>
        internal static void InspectCustomProperties(OpenXmlPackage package, OfficeHiddenDataCollector collector, Action<string>? removeFromModel) {
            CustomFilePropertiesPart? part = GetCustomPropertiesPart(package);
            Properties? properties = part?.Properties;
            if (part == null || properties == null) return;
            foreach (CustomDocumentProperty property in properties.Elements<CustomDocumentProperty>().ToList()) {
                string name = property.Name?.Value ?? string.Empty;
                collector.Add(OfficeHiddenDataKind.CustomProperties, CustomPropertiesLocation, "Custom property '" + name + "'",
                    property.InnerText, () => {
                        property.Remove();
                        removeFromModel?.Invoke(name);
                    });
            }

            if (collector.Removes(OfficeHiddenDataKind.CustomProperties) && !properties.Elements<CustomDocumentProperty>().Any()) {
                package.DeletePart(part);
            }
        }

        internal static void InspectThumbnail(OpenXmlPackage package, OfficeHiddenDataCollector collector) {
            ThumbnailPart? thumbnail = package switch {
                WordprocessingDocument word => word.ThumbnailPart,
                SpreadsheetDocument spreadsheet => spreadsheet.ThumbnailPart,
                PresentationDocument presentation => presentation.ThumbnailPart,
                _ => null
            };
            if (thumbnail == null) return;
            collector.Add(OfficeHiddenDataKind.Thumbnail, thumbnail.Uri.ToString(), "Document thumbnail image", null,
                () => package.DeletePart(thumbnail));
        }

        /// <summary>Custom XML data parts. <paramref name="retainedReason"/> returns why a part must stay, such as a content-control binding.</summary>
        internal static void InspectCustomXml(OpenXmlPartContainer owner, IEnumerable<CustomXmlPart> parts,
            OfficeHiddenDataCollector collector, Func<CustomXmlPart, string?>? retainedReason = null) {
            foreach (CustomXmlPart part in parts.ToList()) {
                string? rootNamespace = ReadRootNamespace(part);
                string? reason = retainedReason?.Invoke(part);
                collector.Add(OfficeHiddenDataKind.CustomXml, part.Uri.ToString(), "Custom XML data part", rootNamespace,
                    reason == null ? () => owner.DeletePart(part) : null, reason);
            }
        }

        /// <summary>The datastore item ID of a custom XML part, used to match content-control bindings.</summary>
        internal static string? GetDataStoreItemId(CustomXmlPart part) {
            CustomXmlPropertiesPart? properties = part.CustomXmlPropertiesPart;
            return properties?.DataStoreItem?.ItemId?.Value;
        }

        /// <summary>Exif, GPS, XMP and IPTC metadata in embedded images, stripped in place without re-encoding pixels.</summary>
        internal static void InspectImageMetadata(OpenXmlPackage package, OfficeHiddenDataCollector collector) {
            foreach (ImagePart image in EnumerateParts(package).Select(pair => pair.Part).OfType<ImagePart>().Distinct().ToList()) {
                byte[] bytes;
                using (Stream stream = image.GetStream(FileMode.Open, FileAccess.Read)) {
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    bytes = buffer.ToArray();
                }

                OfficeImageMetadata metadata;
                try {
                    metadata = OfficeImageMetadata.Read(bytes);
                } catch (Exception exception) when (exception is FormatException or NotSupportedException or ArgumentException) {
                    continue;
                }

                string? description = DescribePrivateImageMetadata(metadata);
                if (description == null) continue;
                byte[]? stripped = null;
                if (collector.Removes(OfficeHiddenDataKind.ImageMetadata)) {
                    try {
                        stripped = StripPrivateImageMetadata(bytes, metadata);
                    } catch (Exception exception) when (exception is FormatException or NotSupportedException or ArgumentException) {
                        stripped = null;
                    }
                }

                collector.Add(OfficeHiddenDataKind.ImageMetadata, image.Uri.ToString(), description, null,
                    stripped == null ? null : () => image.FeedData(new MemoryStream(stripped, writable: false)),
                    "This image container does not support removing metadata without re-encoding.");
            }
        }

        /// <summary>Embedded printer settings parts and the elements that point to them.</summary>
        internal static void InspectPrinterSettings(OpenXmlPackage package, OfficeHiddenDataCollector collector) {
            foreach (PartPair pair in EnumerateParts(package).Where(pair => IsPrinterSettings(pair.Part)).ToList()) {
                collector.Add(OfficeHiddenDataKind.PrinterSettings, pair.Part.Uri.ToString(), "Printer settings", null, () => {
                    if (pair.Parent is OpenXmlPart parentPart) RemoveRelationshipReferences(parentPart, pair.RelationshipId, "printerSettings");
                    pair.Parent.DeletePart(pair.RelationshipId);
                });
            }
        }

        /// <summary>Removes attributes, or whole elements with <paramref name="elementLocalName"/>, that reference a relationship ID.</summary>
        internal static void RemoveRelationshipReferences(OpenXmlPart part, string relationshipId, string elementLocalName) {
            OpenXmlPartRootElement? root = part.RootElement;
            if (root == null) return;
            const string relationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            foreach (OpenXmlElement element in root.Descendants().ToList()) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (attribute.NamespaceUri != relationshipsNamespace || attribute.Value != relationshipId) continue;
                    if (element.LocalName == elementLocalName) element.Remove();
                    else element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                    break;
                }
            }
        }

        internal readonly struct PartPair {
            internal PartPair(OpenXmlPartContainer parent, string relationshipId, OpenXmlPart part) {
                Parent = parent;
                RelationshipId = relationshipId;
                Part = part;
            }

            internal OpenXmlPartContainer Parent { get; }
            internal string RelationshipId { get; }
            internal OpenXmlPart Part { get; }
        }

        /// <summary>Every part relationship in the package, breadth first; a shared part appears once per parent.</summary>
        internal static IEnumerable<PartPair> EnumerateParts(OpenXmlPackage package) {
            var visited = new HashSet<OpenXmlPart>();
            var queue = new Queue<OpenXmlPartContainer>();
            queue.Enqueue(package);
            while (queue.Count > 0) {
                OpenXmlPartContainer container = queue.Dequeue();
                foreach (IdPartPair pair in container.Parts) {
                    yield return new PartPair(container, pair.RelationshipId, pair.OpenXmlPart);
                    if (visited.Add(pair.OpenXmlPart)) queue.Enqueue(pair.OpenXmlPart);
                }
            }
        }

        private static bool IsPrinterSettings(OpenXmlPart part) =>
            part.ContentType.EndsWith(".printerSettings", StringComparison.OrdinalIgnoreCase);

        private static string? DescribePrivateImageMetadata(OfficeImageMetadata metadata) {
            var profiles = new List<string>();
            bool hasGps = false;
            bool hasExif = false;
            foreach (OfficeExifValue value in metadata.ExifValues) {
                if (value.Tag.Directory == OfficeExifDirectory.Gps) hasGps = true;
                else if (!IsOrientation(value.Tag)) hasExif = true;
            }

            if (hasExif) profiles.Add("Exif");
            if (hasGps) profiles.Add("GPS location");
            if (metadata.XmpProfile != null) profiles.Add("XMP");
            if (metadata.IptcProfile != null) profiles.Add("IPTC");
            return profiles.Count == 0 ? null : "Image metadata: " + string.Join(", ", profiles);
        }

        /// <summary>Rewrites only the metadata containers; ICC colour, density and orientation stay so the image looks the same.</summary>
        private static byte[] StripPrivateImageMetadata(byte[] bytes, OfficeImageMetadata metadata) {
            OfficeExifValue? orientation = metadata.GetExifValue(OfficeExifTag.Orientation);
            OfficeImageMetadata stripped = metadata.Clone();
            stripped.ClearExif();
            stripped.XmpProfile = null;
            stripped.IptcProfile = null;
            if (orientation != null) stripped.SetExifValue(OfficeExifTag.Orientation, orientation.Value);
            return OfficeImageMetadata.Apply(bytes, stripped);
        }

        private static bool IsOrientation(OfficeExifTag tag) => tag.Equals(OfficeExifTag.Orientation);

        private static Ext.Properties? GetExtendedProperties(OpenXmlPackage package) => package switch {
            WordprocessingDocument word => word.ExtendedFilePropertiesPart?.Properties,
            SpreadsheetDocument spreadsheet => spreadsheet.ExtendedFilePropertiesPart?.Properties,
            PresentationDocument presentation => presentation.ExtendedFilePropertiesPart?.Properties,
            _ => null
        };

        private static CustomFilePropertiesPart? GetCustomPropertiesPart(OpenXmlPackage package) => package switch {
            WordprocessingDocument word => word.CustomFilePropertiesPart,
            SpreadsheetDocument spreadsheet => spreadsheet.CustomFilePropertiesPart,
            PresentationDocument presentation => presentation.CustomFilePropertiesPart,
            _ => null
        };

        internal static string? ReadRootNamespace(CustomXmlPart part) {
            try {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                    DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                    XmlResolver = null
                });
                return reader.MoveToContent() == System.Xml.XmlNodeType.Element && reader.NamespaceURI.Length > 0 ? reader.NamespaceURI : null;
            } catch (System.Xml.XmlException) {
                return null;
            }
        }
    }
}
