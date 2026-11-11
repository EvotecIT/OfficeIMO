using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
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

        private const string PackageManifestLocation = "/META-INF/content_credential.c2pa";
        private const string UnsupportedImageNote = "This image container does not support removing metadata without re-encoding.";
        private const string BoundImageNote = "Removing this image's metadata would invalidate its Content Credentials; select Provenance to remove both.";
        private const string KeptSourceTypeNote = "The image's digital source type declaration was kept; select Provenance to remove it.";
        private const string PackageManifestNote = "Removed although Provenance was not selected: Content Credentials that cover the whole package cannot stay valid once a loaded document is changed and saved. Use RemoveProvenance on the file to change only provenance.";
        private const string ProvenanceEngineNote = "The provenance engine keeps malformed carriers and source declarations that do not disclose AI generation.";

        /// <summary>
        /// Provenance in the package and its images, and the Exif, GPS, XMP and IPTC metadata in those images. Metadata is stripped in place without
        /// re-encoding pixels. Provenance carriers are found and removed by the provenance engine, so this walk and <c>RemoveProvenance</c> agree.
        /// </summary>
        internal static void InspectImagesAndProvenance(OpenXmlPackage package, OfficeHiddenDataCollector collector) {
            InspectPackageProvenance(package, collector);
            bool removeProvenance = collector.Removes(OfficeHiddenDataKind.Provenance);
            foreach (ImagePart image in EnumerateParts(package).Select(pair => pair.Part).OfType<ImagePart>().Distinct().ToList()) {
                string location = image.Uri.ToString();
                byte[] original = ReadPart(image);
                List<OfficeProvenanceEvidence> provenance = InspectImageProvenance(original, location);
                string? description = DescribePrivateImageMetadata(original);
                if (provenance.Count == 0 && description == null) continue;

                byte[] current = original;
                if (removeProvenance && provenance.Count > 0) current = TryRemoveImageProvenance(current, location) ?? current;

                bool metadataRemoved = false;
                string? metadataNote = null;
                if (description != null && collector.Removes(OfficeHiddenDataKind.ImageMetadata)) {
                    List<OfficeProvenanceEvidence> currentProvenance = ReferenceEquals(current, original) ? provenance : InspectImageProvenance(current, location);
                    bool bound = currentProvenance.Any(IsContentBinding);
                    byte[]? stripped = bound ? null : TryStripPrivateImageMetadata(current, keepDigitalSourceType: !removeProvenance);
                    if (stripped == null) {
                        metadataNote = bound ? BoundImageNote : UnsupportedImageNote;
                    } else {
                        current = stripped;
                        metadataRemoved = true;
                        if (!removeProvenance && provenance.Any(item => item.Carrier == OfficeProvenanceCarrierKind.IptcDigitalSourceType)) metadataNote = KeptSourceTypeNote;
                    }
                }

                if (!ReferenceEquals(current, original)) image.FeedData(new MemoryStream(current, writable: false));

                if (description != null) {
                    OfficeHiddenDataOutcome outcome = collector.Resolve(OfficeHiddenDataKind.ImageMetadata, metadataRemoved, metadataNote, out string? note);
                    collector.Record(OfficeHiddenDataKind.ImageMetadata, location, description, null, outcome, metadataRemoved ? metadataNote : note);
                }

                List<OfficeProvenanceEvidence> remaining = ReferenceEquals(current, original) ? provenance.ToList() : InspectImageProvenance(current, location);
                foreach (OfficeProvenanceEvidence evidence in provenance) {
                    bool removed = !TakeMatch(remaining, evidence);
                    OfficeHiddenDataOutcome outcome = collector.Resolve(OfficeHiddenDataKind.Provenance, removed, ProvenanceEngineNote, out string? note);
                    collector.Record(OfficeHiddenDataKind.Provenance, location + "/" + evidence.Location, DescribeProvenance(evidence),
                        ProvenanceValue(evidence), outcome, note, evidence);
                }
            }
        }

        /// <summary>
        /// A C2PA manifest stored as a package part. It binds every byte of the package, so no edit of a loaded document can keep it valid,
        /// and saving does not carry an unreferenced part. Removal therefore always deletes it and says so.
        /// </summary>
        private static void InspectPackageProvenance(OpenXmlPackage package, OfficeHiddenDataCollector collector) {
#pragma warning disable OOXML0001
            IPackage opc = package.GetPackage();
            var uri = new Uri(PackageManifestLocation, UriKind.Relative);
            if (!opc.PartExists(uri)) return;
            byte[] manifest;
            using (Stream stream = opc.GetPart(uri).GetStream(FileMode.Open, FileAccess.Read)) manifest = ReadAll(stream);
            var limits = new OfficeProvenanceOptions();
            bool valid = manifest.LongLength <= limits.MaxManifestBytes &&
                OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, limits.MaxManifestBytes, limits.MaxContainerEntries, out _);
            OfficeProvenanceEvidence evidence = new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest,
                "ZIP/META-INF/content_credential.c2pa[0]", valid, manifest.Length)
                .WithManifest(valid ? OfficeC2paManifestStore.TryDescribe(manifest, 0, manifest.Length) : null);

            bool remove = collector.Removal != null;
            if (remove) {
                IdPartPair related = package.Parts.FirstOrDefault(pair => pair.OpenXmlPart.Uri.OriginalString == PackageManifestLocation);
                if (related.OpenXmlPart != null) {
                    package.DeletePart(related.RelationshipId);
                } else {
                    foreach (IPackageRelationship relationship in opc.Relationships.ToList()) {
                        if (relationship.TargetMode == TargetMode.Internal &&
                            relationship.TargetUri.OriginalString.TrimStart('/') == PackageManifestLocation.TrimStart('/')) {
                            opc.Relationships.Remove(relationship.Id);
                        }
                    }
                }
                if (opc.PartExists(uri)) opc.DeletePart(uri);
            }
#pragma warning restore OOXML0001

            OfficeHiddenDataOutcome outcome = remove ? OfficeHiddenDataOutcome.Removed : OfficeHiddenDataOutcome.Found;
            string? note = remove && !collector.Removes(OfficeHiddenDataKind.Provenance) ? PackageManifestNote : null;
            collector.Record(OfficeHiddenDataKind.Provenance, PackageManifestLocation, DescribeProvenance(evidence),
                ProvenanceValue(evidence), outcome, note, evidence);
        }

        private static List<OfficeProvenanceEvidence> InspectImageProvenance(byte[] data, string location) {
            var options = new OfficeProvenanceOptions { ProcessEmbeddedAssets = false };
            try {
                if (data.LongLength > options.MaxAssetBytes || !IsProvenanceImage(OfficeProvenanceInspector.DetectFormat(data, null, options))) {
                    return new List<OfficeProvenanceEvidence>();
                }
                return OfficeProvenanceInspector.Inspect(data, location, options).Evidence.ToList();
            } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or NotSupportedException or FormatException) {
                return new List<OfficeProvenanceEvidence>();
            }
        }

        private static byte[]? TryRemoveImageProvenance(byte[] data, string location) {
            try {
                OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(data, location, new OfficeProvenanceRemovalOptions { ProcessEmbeddedAssets = false });
                return result.WasChanged ? result.ToArray() : null;
            } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or NotSupportedException or FormatException) {
                return null;
            }
        }

        private static bool IsProvenanceImage(OfficeProvenanceAssetFormat format) =>
            format is OfficeProvenanceAssetFormat.Jpeg or OfficeProvenanceAssetFormat.Png or OfficeProvenanceAssetFormat.Webp or
                OfficeProvenanceAssetFormat.Gif or OfficeProvenanceAssetFormat.Tiff or OfficeProvenanceAssetFormat.Svg;

        /// <summary>Carriers whose signature covers the image bytes, so any metadata edit breaks them.</summary>
        private static bool IsContentBinding(OfficeProvenanceEvidence evidence) =>
            evidence.Carrier is OfficeProvenanceCarrierKind.C2paManifest or OfficeProvenanceCarrierKind.C2paExternalManifest;

        private static bool TakeMatch(List<OfficeProvenanceEvidence> remaining, OfficeProvenanceEvidence evidence) {
            int index = remaining.FindIndex(item => item.Carrier == evidence.Carrier && item.Value == evidence.Value &&
                item.PayloadLength == evidence.PayloadLength);
            if (index < 0) return false;
            remaining.RemoveAt(index);
            return true;
        }

        private static string DescribeProvenance(OfficeProvenanceEvidence evidence) {
            string description = evidence.Carrier switch {
                OfficeProvenanceCarrierKind.C2paManifest => evidence.Manifest?.DeclaresGenerativeAi == true
                    ? "Content Credentials (C2PA manifest) declaring generative AI"
                    : "Content Credentials (C2PA manifest)",
                OfficeProvenanceCarrierKind.C2paExternalManifest => "Reference to external Content Credentials",
                _ => "Digital source type: " + evidence.DigitalSourceKind
            };
            return evidence.IsStructurallyValid ? description : description + " (malformed)";
        }

        private static string? ProvenanceValue(OfficeProvenanceEvidence evidence) =>
            evidence.Carrier == OfficeProvenanceCarrierKind.C2paManifest ? evidence.Manifest?.ClaimGenerator : evidence.Value;

        private static byte[] ReadPart(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return ReadAll(stream);
        }

        private static byte[] ReadAll(Stream stream) {
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
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

        private static string? DescribePrivateImageMetadata(byte[] data) {
            OfficeImageMetadata metadata;
            try {
                metadata = OfficeImageMetadata.Read(data);
            } catch (Exception exception) when (exception is FormatException or NotSupportedException or ArgumentException) {
                return null;
            }

            var profiles = new List<string>();
            bool hasGps = false;
            bool hasExif = false;
            foreach (OfficeExifValue value in metadata.ExifValues) {
                if (value.Tag.Directory == OfficeExifDirectory.Gps) hasGps = true;
                else if (!IsOrientation(value.Tag)) hasExif = true;
            }

            if (hasExif) profiles.Add("Exif");
            if (hasGps) profiles.Add("GPS location");
            byte[]? xmp = metadata.XmpProfileData;
            if (xmp != null && !OfficeProvenanceXmp.HasOnlyDigitalSourceTypes(xmp, new OfficeProvenanceOptions())) profiles.Add("XMP");
            if (metadata.IptcProfile != null) profiles.Add("IPTC");
            return profiles.Count == 0 ? null : "Image metadata: " + string.Join(", ", profiles);
        }

        /// <summary>
        /// Rewrites only the metadata containers; ICC colour, density and orientation stay so the image looks the same.
        /// With <paramref name="keepDigitalSourceType"/>, XMP is reduced to the image's digital source type declaration.
        /// </summary>
        private static byte[]? TryStripPrivateImageMetadata(byte[] data, bool keepDigitalSourceType) {
            try {
                OfficeImageMetadata metadata = OfficeImageMetadata.Read(data);
                OfficeExifValue? orientation = metadata.GetExifValue(OfficeExifTag.Orientation);
                OfficeImageMetadata stripped = metadata.Clone();
                stripped.ClearExif();
                byte[]? xmp = metadata.XmpProfileData;
                stripped.XmpProfile = keepDigitalSourceType && xmp != null
                    ? OfficeProvenanceXmp.CreateDigitalSourceTypePacket(xmp, new OfficeProvenanceOptions())
                    : null;
                stripped.IptcProfile = null;
                if (orientation != null) stripped.SetExifValue(OfficeExifTag.Orientation, orientation.Value);
                return OfficeImageMetadata.Apply(data, stripped);
            } catch (Exception exception) when (exception is FormatException or NotSupportedException or ArgumentException) {
                return null;
            }
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
