using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private sealed record PackageSaveProjection(string Main, string? DestinationContentType, string? RootRelationships, string? MainRelationshipId,
            Dictionary<string, OfficeOpcEditablePackage.OutputPart> Parts);

        private bool TryPlanPackageSaveProjection(OfficeOpcEditablePackage package, string mainTarget, string? destinationContentType,
            bool stripVba, out PackageSaveProjection? projection) {
            projection = null;
            try {
                if (package.GetPartName(mainTarget) is not { } main) return false;
                var parts = new Dictionary<string, OfficeOpcEditablePackage.OutputPart>(StringComparer.Ordinal);
                string rootRelationships = OfficeOpcPartNames.GetRelationshipsPartName(null);
                string? rootPart = package.GetPartName(rootRelationships), mainId = null;
                if (rootPart == null || ReadEditablePackageRelationships(package, null, strictXml: true) is not { } packageLinks) return false;
                if (destinationContentType != null) {
                    if (main.IndexOf('%') >= 0 || main.Any(char.IsWhiteSpace)) return false;
                    string destination = NextSubtypeMainPartName(package, main, out _);
                    OfficeOpcRelationship[] mainLinks = packageLinks.Where(edge => !edge.IsExternal &&
                        edge.Type == OfficeOpcPackage.OfficeDocumentRelationshipType).ToArray();
                    if (mainLinks.Length != 1) return false;
                    mainId = mainLinks[0].Id;
                    parts.Add(main, new OfficeOpcEditablePackage.OutputPart(destination));
                } else if (!packageLinks.Any(edge => edge.IsExternal)) {
                    rootPart = null;
                }

                string? mainRelationships = package.GetPartName(OfficeOpcPartNames.GetRelationshipsPartName(main));
                if (mainRelationships != null && destinationContentType != null) {
                    string renamedRelationships = OfficeOpcPartNames.GetRelationshipsPartName(parts[main].Name);
                    parts.Add(mainRelationships, new OfficeOpcEditablePackage.OutputPart(renamedRelationships));
                }

                if (stripVba) {
                    if (mainRelationships == null || !TryPlanPackageVbaMutation(package, remove: true, out PackageVbaMutation? vba) ||
                        vba?.Relationships == null) return false;
                    string? vbaId = vba.Relationships.GetAttribute(vba.RemovedRelationships[0], OfficeXmlNamespaces.None, "Id");
                    if (vbaId == null) return false;
                    OfficeXmlTree remaining = CopyPackageTree(package, mainRelationships);
                    int projectEdge = FindRelationship(remaining, vbaId);
                    if (projectEdge == OfficeXmlTree.NoNode) return false;
                    remaining.Remove(projectEdge);
                    string savedRelationships = destinationContentType == null ? mainRelationships : parts[mainRelationships].Name!;
                    parts[mainRelationships] = remaining.FindChild(remaining.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship") == OfficeXmlTree.NoNode
                        ? default : new OfficeOpcEditablePackage.OutputPart(savedRelationships, remaining);
                    foreach (string deleted in vba.RemovedParts) {
                        parts.Add(deleted, default);
                        if (package.GetPartName(OfficeOpcPartNames.GetRelationshipsPartName(deleted)) is { } deletedRelationships) {
                            parts.Add(deletedRelationships, default);
                        }
                    }
                }

                projection = new PackageSaveProjection(main, destinationContentType, rootPart, mainId, parts);
                return true;
            } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or IOException
                or ArgumentException or NotSupportedException or InvalidOperationException or OverflowException) {
                return false;
            }
        }

        private static string NextSubtypeMainPartName(OfficeOpcEditablePackage package, string main, out bool hadCollision) {
            int dot = main.LastIndexOf('.');
            bool hasExtension = dot > main.LastIndexOf('/');
            string stem = hasExtension ? main.Substring(0, dot) : main;
            string extension = hasExtension ? main.Substring(dot) : string.Empty;
            int suffix = 2;
            string destination;
            do {
                destination = stem + suffix.ToString(CultureInfo.InvariantCulture) + extension;
                suffix++;
            } while (package.PartExists(destination) || package.PartExists(OfficeOpcPartNames.GetRelationshipsPartName(destination)));
            hadCollision = suffix > 3;
            return destination;
        }

        private static OfficeXmlTree CopyPackageTree(OfficeOpcEditablePackage package, string name) {
            byte[] content = package.ReadPartBytes(name);
            return OfficeXmlTree.Parse(content, content.Length, package.Namespaces);
        }

        private static IReadOnlyDictionary<string, OfficeOpcEditablePackage.OutputPart> CreatePackageSaveProjection(
            OfficeOpcEditablePackage package, PackageSaveProjection projection) {
            if (projection.RootRelationships is { } rootPart) {
                OfficeXmlTree root = CopyPackageTree(package, rootPart);
                if (projection.MainRelationshipId is { } id) {
                    int mainRelationship = FindRelationship(root, id);
                    root.SetAttribute(mainRelationship, OfficeXmlNamespaces.None, "Target", projection.Parts[projection.Main].Name!);
                }
                const int rel = OfficeXmlNamespaces.PackageRelationships;
                for (int relationship = root.FindChild(root.Root, rel, "Relationship"), next; relationship != OfficeXmlTree.NoNode; relationship = next) {
                    next = root.FindNextSibling(relationship, rel, "Relationship");
                    if (root.GetAttribute(relationship, OfficeXmlNamespaces.None, "TargetMode") == "External") root.Remove(relationship);
                }
                projection.Parts.Add(rootPart, new OfficeOpcEditablePackage.OutputPart(rootPart, root));
            }
            string catalogPart = package.GetPartName(OfficeOpcEditablePackage.ContentTypesPartName)!;
            OfficeXmlTree types = CopyPackageTree(package, catalogPart);
            const int ns = OfficeXmlNamespaces.ContentTypes;
            const int attribute = OfficeXmlNamespaces.None;
            bool mainOverride = false;
            for (int node = types.FindChild(types.Root, ns, "Override"), next; node != OfficeXmlTree.NoNode; node = next) {
                next = types.FindNextSibling(node, ns, "Override");
                if (types.GetAttribute(node, attribute, "PartName") is not { } declared) continue;
                string? part = package.GetPartName(declared);
                if (part == null) {
                    types.Remove(node);
                    continue;
                }
                if (!projection.Parts.TryGetValue(part, out OfficeOpcEditablePackage.OutputPart output) || part == output.Name) continue;
                if (output.Name == null) {
                    types.Remove(node);
                    continue;
                }
                types.SetAttribute(node, attribute, "PartName", output.Name);
                if (part == projection.Main && projection.DestinationContentType is { } newType) {
                    types.SetAttribute(node, attribute, "ContentType", newType);
                    mainOverride = true;
                }
            }
            if (projection.DestinationContentType is { } contentType && !mainOverride) {
                int added = types.CreateElement(ns, "Override");
                types.SetAttribute(added, attribute, "PartName", projection.Parts[projection.Main].Name!);
                types.SetAttribute(added, attribute, "ContentType", contentType);
                types.Append(types.Root, added);
            }
            projection.Parts.Add(catalogPart, new OfficeOpcEditablePackage.OutputPart(catalogPart, types));
            return projection.Parts;
        }
    }
}
