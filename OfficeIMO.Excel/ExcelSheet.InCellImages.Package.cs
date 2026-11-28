using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly record struct RichImageLocation(int Future, int Value, int Structure, int Relationship, string AltText);

        private sealed class RichImageIndex {
            private readonly OfficeXmlNode[] _types;
            private readonly Func<int, OfficeXmlNode> _metadataBlock;
            private readonly Func<int, OfficeXmlNode> _futureBlock;
            private readonly Func<int, OfficeXmlNode> _value;
            private readonly Func<int, OfficeXmlNode> _structure;

            internal RichImageIndex(OfficeXmlNode metadata, Func<int, OfficeXmlNode> metadataBlock,
                Func<int, OfficeXmlNode> futureBlock, Func<int, OfficeXmlNode> value, Func<int, OfficeXmlNode> structure) {
                _types = metadata.Child(WorksheetMarkup, "metadataTypes").Children(WorksheetMarkup, "metadataType").ToList().ToArray();
                _metadataBlock = metadataBlock;
                _futureBlock = futureBlock;
                _value = value;
                _structure = structure;
            }

            internal bool TryResolve(uint metadataIndex, out RichImageLocation image) {
                image = default;
                if (metadataIndex == 0 || metadataIndex > int.MaxValue) return false;

                OfficeXmlNode block = _metadataBlock((int)metadataIndex - 1);
                foreach (OfficeXmlNode record in block.Children(WorksheetMarkup, "rc")) {
                    uint type = record.UInt32Attribute("t") ?? 0U;
                    if (type == 0 || type > _types.Length ||
                        !string.Equals(_types[type - 1].Attribute("name"), RichValueMetadataName, StringComparison.OrdinalIgnoreCase)) continue;

                    if (record.Int32Attribute("v") is not int futureIndex || futureIndex < 0
                        || RichValueErrorLookup.FindValueBlock(_futureBlock(futureIndex)).Int32Attribute("i") is not int valueIndex || valueIndex < 0) {
                        return false;
                    }

                    OfficeXmlNode value = _value(valueIndex);
                    if (value.Int32Attribute("s") is not int structureIndex || structureIndex < 0
                        || !TryGetImageRelationshipIndex(value, _structure(structureIndex), out int relationshipIndex, out string altText)) return false;
                    image = new RichImageLocation(futureIndex, valueIndex, structureIndex, relationshipIndex, altText);
                    return true;
                }
                return false;
            }
        }

        private readonly record struct PackageRichImage(string Reference, string PartName, string ContentType, string AltText);

        private static IReadOnlyList<PackageRichImage> ReadOwnedInCellImageLocations(OfficeOpcPackage package, ExcelDeferredSheet sheet) {
            const string rich = RichValueErrorLookup.RichData;
            const string richRelationships = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
            string? workbook = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
            if (workbook == null) return Array.Empty<PackageRichImage>();

            string? metadataPart = package.GetRelatedPartName(workbook, ExcelReadWorkbookParts.Metadata.Relationship);
            string? valuePart = package.GetRelatedPartName(workbook, ExcelReadWorkbookParts.RichValues.Relationship);
            string? structurePart = package.GetRelatedPartName(workbook, ExcelReadWorkbookParts.RichStructures.Relationship);
            string? relationshipPart = package.GetRelatedPartName(workbook, RichValueRelRelationshipType);
            if (metadataPart == null || valuePart == null || structurePart == null || relationshipPart == null) return Array.Empty<PackageRichImage>();
            if (!package.PartExists(metadataPart) || !package.PartExists(valuePart)
                || !package.PartExists(structurePart) || !package.PartExists(relationshipPart)) return Array.Empty<PackageRichImage>();

            using OfficePooledBytes metadataBytes = ReadImageMetadata(package, metadataPart, MaximumRichValueMetadataBytes, "Cell metadata");
            using OfficePooledBytes valueBytes = ReadImageMetadata(package, valuePart, MaximumRichValueMetadataBytes, "Rich-value data");
            using OfficePooledBytes structureBytes = ReadImageMetadata(package, structurePart, MaximumRichValueMetadataBytes, "Rich-value structures");
            using OfficePooledBytes relationshipBytes = ReadImageMetadata(package, relationshipPart, MaximumRichValueRelationshipBytes, "Rich-value relationship metadata");
            OfficeXmlNode Root(OfficePooledBytes bytes) {
                OfficeXmlTree tree = OfficeXmlTree.Parse(bytes.Array, bytes.Length, package.Namespaces);
                return new OfficeXmlNode(tree, tree.Root);
            }

            OfficeXmlNode metadata = Root(metadataBytes);
            OfficeXmlNode future = default;
            foreach (OfficeXmlNode candidate in metadata.Children(WorksheetMarkup, "futureMetadata")) {
                if (!string.Equals(candidate.Attribute("name"), RichValueMetadataName, StringComparison.OrdinalIgnoreCase)) continue;
                future = candidate;
                break;
            }
            if (future.IsNone) return Array.Empty<PackageRichImage>();

            var blocks = metadata.Child(WorksheetMarkup, "valueMetadata").Children(WorksheetMarkup, "bk").ToList();
            var futureBlocks = future.Children(WorksheetMarkup, "bk").ToList();
            var values = Root(valueBytes).Children().ToList();
            values.RemoveAll(value => !value.Is(rich, "rv"));
            var structures = Root(structureBytes).Children().ToList();
            structures.RemoveAll(structure => !structure.Is(rich, "s"));
            var relationships = Root(relationshipBytes).Children().ToList();
            relationships.RemoveAll(relationship => !relationship.Is(richRelationships, "rel"));
            var index = new RichImageIndex(metadata, blocks.ElementAtOrDefault, futureBlocks.ElementAtOrDefault,
                values.ElementAtOrDefault, structures.ElementAtOrDefault);

            var images = new List<PackageRichImage>();
            var resolved = new Dictionary<uint, (string PartName, string ContentType, string AltText)?>();
            OfficeXmlPullReader reader = package.OpenXml(sheet.PartName, out OfficePooledBytes worksheet, processMarkupCompatibility: false);
            using (worksheet) {
                bool sheetData = false;
                bool row = false;
                while (reader.Read()) {
                    if (reader.Depth == 2) {
                        if (reader.NodeKind == OfficeXmlNodeKind.EndElement && sheetData) break;
                        if (reader.NodeKind == OfficeXmlNodeKind.StartElement) sheetData = reader.IsElement(WorksheetMarkup, "sheetData");
                    }
                    if (sheetData && reader.Depth == 3) {
                        row = reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.IsElement(WorksheetMarkup, "row");
                    }
                    if (!row || reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 4
                        || !reader.IsElement(WorksheetMarkup, "c")
                        || !reader.TryGetAttribute(OfficeXmlNamespaces.None, "vm", out string metadataIndex)
                        || !uint.TryParse(metadataIndex, out uint id)) continue;

                    if (!resolved.TryGetValue(id, out var image)) {
                        image = ReadImage(id);
                        resolved.Add(id, image);
                    }
                    if (image is { } available) {
                        reader.TryGetAttribute(OfficeXmlNamespaces.None, "r", out string reference);
                        images.Add(new PackageRichImage(reference, available.PartName, available.ContentType, available.AltText));
                    }
                }
            }
            return images;

            (string PartName, string ContentType, string AltText)? ReadImage(uint id) {
                if (!index.TryResolve(id, out RichImageLocation image) || image.Relationship >= relationships.Count
                    || relationships[image.Relationship].Attribute(OfficeXmlNamespaces.Relationships, "id") is not string relationship
                    || package.FindRelationship(relationshipPart, relationship)?.TargetPartName is not string target
                    || !package.PartExists(target)
                    || package.GetContentType(target) is not string contentType
                    || !contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return null;
                return (target, contentType, image.AltText);
            }
        }

        private static OfficePooledBytes ReadImageMetadata(OfficeOpcPackage package, string part, long maximum, string description) {
            if (package.GetPartLength(part) > maximum) {
                throw new InvalidDataException($"{description} exceeds the supported {maximum}-byte limit.");
            }
            return package.ReadPart(part);
        }

        private static byte[] ReadOwnedImagePayload(OfficeOpcPackage package, string partName, long remaining, long maximum) {
            if (package.GetPartLength(partName) > remaining) {
                throw new InvalidOperationException($"In-cell image payloads exceed maximumTotalImageBytes ({maximum}).");
            }
            using OfficePooledBytes bytes = package.ReadPart(partName);
            using var stream = new MemoryStream(bytes.Array, 0, bytes.Length, writable: false);
            return ReadInCellImagePayload(stream, remaining, maximum);
        }
    }
}
