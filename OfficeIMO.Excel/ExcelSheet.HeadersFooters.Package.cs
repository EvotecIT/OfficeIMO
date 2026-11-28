using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string HeaderFooterVmlContentType = "application/vnd.openxmlformats-officedocument.vmlDrawing";

        private static (Dictionary<string, ExcelHeaderFooterImageInfo> Images, bool HasDrawingPart) ReadOwnedHeaderFooterImages(
            OfficeOpcPackage package, ExcelDeferredSheet sheet, long maximumTotalSourceImageBytes) {
            var empty = new Dictionary<string, ExcelHeaderFooterImageInfo>(StringComparer.OrdinalIgnoreCase);
            OfficeXmlPullReader reader = package.OpenXml(sheet.PartName, out OfficePooledBytes sheetBytes, processMarkupCompatibility: false);
            string? relationshipId = null;
            using (sheetBytes) {
                while (reader.Read()) {
                    if (reader.Depth != 2 || reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                    if (reader.IsElement(WorksheetMarkup, "legacyDrawingHF")) {
                        if (reader.TryGetAttribute(OfficeXmlNamespaces.Relationships, "id", out string id)) relationshipId = id;
                        break;
                    }
                    if (!reader.IsEmptyElement) reader.SkipElement();
                }
            }
            string? drawing = relationshipId == null ? null : package.FindRelationship(sheet.PartName, relationshipId)?.TargetPartName;
            if (drawing == null || package.GetContentType(drawing) != HeaderFooterVmlContentType || !package.PartExists(drawing)) {
                return (empty, false);
            }
            if (package.GetPartLength(drawing) > MaxVmlXmlPartBytes) return (empty, true);
            using OfficePooledBytes vml = package.ReadPart(drawing);
            using var source = new MemoryStream(vml.Array, 0, vml.Length, writable: false);
            return (ReadVmlHeaderFooterImages(source, maximumTotalSourceImageBytes, (id, limit) => {
                string? target = package.FindRelationship(drawing, id)?.TargetPartName;
                if (target == null || !package.PartExists(target) || package.GetPartLength(target) > limit) return default;
                string? type = package.GetContentType(target);
                if (type?.StartsWith("image/", StringComparison.OrdinalIgnoreCase) != true) return default;
                using OfficePooledBytes content = package.ReadPart(target);
                var bytes = new byte[content.Length];
                Buffer.BlockCopy(content.Array, 0, bytes, 0, bytes.Length);
                return (bytes, type);
            }), true);
        }

        private static void WriteOwnedHeaderFooterPicture(OfficeOpcEditablePackage package, ExcelDeferredSheet sheet, OfficeXmlNode worksheet,
            string shape, byte[] bytes, string contentType, double width, double height) {
            string? relationship = worksheet.Child(WorksheetMarkup, "legacyDrawingHF").Attribute(OfficeXmlNamespaces.Relationships, "id");
            string? drawing = relationship == null ? null : package.GetRelationshipTarget(sheet.PartName, relationship);
            if (drawing != null && package.GetContentType(drawing) != HeaderFooterVmlContentType) {
                throw new InvalidDataException("The header/footer drawing relationship has the wrong content type.");
            }

            bool hasDrawing = drawing != null;
            drawing ??= FindUnusedObjectPartName(package, "/xl/drawings/vmldrawing", ".vml");
            XDocument vml;
            if (hasDrawing && package.TryGetBinaryLength(drawing, out long partLength) && partLength <= MaxVmlXmlPartBytes) {
                using OfficePooledBytes content = package.ReadPart(drawing);
                using var source = new MemoryStream(content.Array, 0, content.Length, writable: false);
                vml = LoadOrCreateHeaderFooterVmlDocument(source);
            } else {
                vml = LoadOrCreateHeaderFooterVmlDocument(null);
            }
            if (!hasDrawing) {
                package.CreateBinaryPart(drawing, HeaderFooterVmlContentType, Array.Empty<byte>());
                relationship = package.AddRelationship(sheet.PartName,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                    OfficeOpcPartNames.GetRelativeTarget(sheet.PartName, drawing));
                SetHeaderFooterDrawingRelationship(worksheet, relationship);
            }

            (string imageContentType, string imageExtension) = ToImagePartContentType(contentType);
            string image = FindUnusedObjectPartName(package, "/xl/media/image", imageExtension.StartsWith(".", StringComparison.Ordinal)
                ? imageExtension : "." + imageExtension);
            package.CreateBinaryPart(image, imageContentType, bytes);
            string imageRelationship = package.AddRelationship(drawing,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                OfficeOpcPartNames.GetRelativeTarget(drawing, image));
            UpsertHeaderFooterVmlShape(vml, shape, imageRelationship, width, height);
            using var output = new MemoryStream();
            vml.Save(output, SaveOptions.DisableFormatting);
            byte[] payload = output.ToArray();
            package.ReplacePart(drawing, payload, payload.Length);
        }

        private static string FindUnusedObjectPartName(OfficeOpcEditablePackage package, string prefix, string extension) {
            string original = prefix + extension;
            if (!package.PartExists(original)) return original;
            for (int number = 1; number < int.MaxValue; number++) {
                string name = prefix + number.ToString(CultureInfo.InvariantCulture) + extension;
                if (!package.PartExists(name)) return name;
            }
            throw new InvalidDataException("No part name is available for the worksheet object.");
        }

        private static void CleanupOwnedHeaderFooterPicture(OfficeOpcEditablePackage package, ExcelDeferredSheet sheet, OfficeXmlNode worksheet) {
            OfficeXmlNode drawing = worksheet.Child(WorksheetMarkup, "legacyDrawingHF");
            if (drawing.Attribute(OfficeXmlNamespaces.Relationships, "id") is not string id) return;
            string? target = package.GetRelationshipTarget(sheet.PartName, id);
            if (target != null && WorksheetHeaderContainsPicture(worksheet)) return;
            drawing.Remove();
            if (target == null || package.GetContentType(target) != HeaderFooterVmlContentType) return;

            OfficeXmlTree links = package.GetRelationships(sheet.PartName);
            var relationships = new OfficeXmlNode(links, links.Root);
            foreach (OfficeXmlNode link in relationships.Children(OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (link.Attribute("Id") != id) continue;
                link.Remove();
                break;
            }
            if (relationships.FirstChild.IsNone) package.DeletePart(OfficeOpcPartNames.GetRelationshipsPartName(sheet.PartName));
            if (HasOwnedPartReference(package, target)) return;

            string vmlRelationships = OfficeOpcPartNames.GetRelationshipsPartName(target);
            if (package.PartExists(vmlRelationships)) {
                OfficeXmlTree vmlLinks = package.GetXml(vmlRelationships);
                var images = new List<string>();
                foreach (OfficeXmlNode link in new OfficeXmlNode(vmlLinks, vmlLinks.Root).Children(OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (link.Attribute("TargetMode") == "External") continue;
                    string? image = OfficeOpcPartNames.ResolveRelationshipTarget(vmlRelationships, link.Attribute("Target"));
                    if (image != null && package.GetContentType(image)?.StartsWith("image/", StringComparison.OrdinalIgnoreCase) == true) {
                        images.Add(image);
                    }
                }
                package.DeletePart(vmlRelationships);
                foreach (string image in images) {
                    if (!HasOwnedPartReference(package, image)) package.DeletePart(image);
                }
            }
            package.DeletePart(target);
        }

        private static bool HasOwnedPartReference(OfficeOpcEditablePackage package, string target) {
            string identity = OfficeOpcPackage.GetPartIdentity(target);
            foreach (string name in package.PartNames) {
                if (!name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                OfficeXmlTree tree = package.GetXml(name);
                foreach (OfficeXmlNode relationship in new OfficeXmlNode(tree, tree.Root).Children(OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (relationship.Attribute("TargetMode") == "External") continue;
                    string? nameTarget = OfficeOpcPartNames.ResolveRelationshipTarget(name, relationship.Attribute("Target"));
                    if (nameTarget != null && OfficeOpcPackage.GetPartIdentity(nameTarget) == identity) return true;
                }
            }
            return false;
        }

        private static bool WorksheetHeaderContainsPicture(OfficeXmlNode worksheet) {
            OfficeXmlNode footer = worksheet.Child(WorksheetMarkup, "headerFooter");
            return HasHeaderFooterCode(GetHeaderFooterText(footer, ExcelHeaderFooterKind.OddHeader), 'G')
                || HasHeaderFooterCode(GetHeaderFooterText(footer, ExcelHeaderFooterKind.OddFooter), 'G')
                || HasHeaderFooterCode(GetHeaderFooterText(footer, ExcelHeaderFooterKind.FirstHeader), 'G')
                || HasHeaderFooterCode(GetHeaderFooterText(footer, ExcelHeaderFooterKind.FirstFooter), 'G')
                || HasHeaderFooterCode(GetHeaderFooterText(footer, ExcelHeaderFooterKind.EvenHeader), 'G')
                || HasHeaderFooterCode(GetHeaderFooterText(footer, ExcelHeaderFooterKind.EvenFooter), 'G');
        }
    }
}
