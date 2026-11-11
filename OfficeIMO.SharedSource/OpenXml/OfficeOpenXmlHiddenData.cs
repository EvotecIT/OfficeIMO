using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using System.IO.Packaging;
using Ext = DocumentFormat.OpenXml.ExtendedProperties;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>Package-level hidden data shared by Word, Excel and PowerPoint: properties, thumbnail, custom XML, provenance, image metadata and printer settings.</summary>
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

        /// <summary>Provenance and image metadata in the package and its images, through the shared asset engine.</summary>
        internal static void InspectImagesAndProvenance(OpenXmlPackage package, OfficeHiddenDataCollector collector) {
            InspectPackageProvenance(package, collector);
            var related = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OpenXmlPart part in EnumerateParts(package).Select(pair => pair.Part).Distinct().ToList()) {
                related.Add(part.Uri.OriginalString);
                if (part is not ImagePart image) continue;
                byte[] original;
                using (Stream stream = image.GetStream(FileMode.Open, FileAccess.Read)) original = OfficeHiddenDataAssets.ReadAll(stream);
                byte[]? replacement = OfficeHiddenDataAssets.ProcessImage(original, image.Uri.ToString(), collector);
                if (replacement != null) image.FeedData(new MemoryStream(replacement, writable: false));
            }
            InspectUnusedImages(package, related, collector);
        }

        /// <summary>
        /// Pictures stored in the package that no relationship reaches. Nothing shows them, but their metadata and provenance still leave with
        /// the file, so they are reported and cleaned like any other picture.
        /// </summary>
        private static void InspectUnusedImages(OpenXmlPackage package, HashSet<string> related, OfficeHiddenDataCollector collector) {
#pragma warning disable OOXML0001
            IPackage opc = package.GetPackage();
            foreach (IPackagePart part in opc.GetParts().ToList()) {
                string uri = part.Uri.OriginalString;
                if (related.Contains(uri) || !part.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) continue;
                byte[] original;
                using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read)) original = OfficeHiddenDataAssets.ReadAll(stream);
                byte[]? replacement = OfficeHiddenDataAssets.ProcessImage(original, uri, collector, unused: true);
                if (replacement == null) continue;
                using Stream target = part.GetStream(FileMode.Create, FileAccess.Write);
                target.Write(replacement, 0, replacement.Length);
            }
#pragma warning restore OOXML0001
        }

        /// <summary>A package-level C2PA manifest part. Saving a loaded package does not carry an unreferenced part, so removal deletes it.</summary>
        private static void InspectPackageProvenance(OpenXmlPackage package, OfficeHiddenDataCollector collector) {
#pragma warning disable OOXML0001
            IPackage opc = package.GetPackage();
            string location = "/" + OfficeHiddenDataAssets.PackageManifestPath;
            var uri = new Uri(location, UriKind.Relative);
            if (!opc.PartExists(uri)) return;
            byte[] manifest;
            using (Stream stream = opc.GetPart(uri).GetStream(FileMode.Open, FileAccess.Read)) manifest = OfficeHiddenDataAssets.ReadAll(stream);
            if (!OfficeHiddenDataAssets.ProcessPackageManifest(manifest, location, collector)) return;
            IdPartPair related = package.Parts.FirstOrDefault(pair => pair.OpenXmlPart.Uri.OriginalString == location);
            if (related.OpenXmlPart != null) {
                package.DeletePart(related.RelationshipId);
            } else {
                foreach (IPackageRelationship relationship in opc.Relationships.ToList()) {
                    if (relationship.TargetMode == TargetMode.Internal &&
                        relationship.TargetUri.OriginalString.TrimStart('/') == OfficeHiddenDataAssets.PackageManifestPath) {
                        opc.Relationships.Remove(relationship.Id);
                    }
                }
            }
            if (opc.PartExists(uri)) opc.DeletePart(uri);
#pragma warning restore OOXML0001
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
