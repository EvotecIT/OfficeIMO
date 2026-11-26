using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Image parts added to the package without the SDK. As through the SDK, an image added again with the same content
    /// shares the part it was stored in.
    /// </summary>
    public sealed partial class PowerPointDocument {
        private const string ImageRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "image";

        // Image parts added without the SDK, by content, with the pixel size read from them.
        private Dictionary<string, (string Part, int Width, int Height)>? _packageImageParts;

        /// <summary>
        /// Links a slide to an image part holding <paramref name="data"/>, adding the part unless an identical image was added
        /// before; false, having changed nothing, when the model has to add it.
        /// </summary>
        internal bool TryAddPackageImagePart(string slidePart, byte[] data, string contentType, string extension,
            out string relationshipId, out int width, out int height) {
            relationshipId = string.Empty;
            width = height = 0;
            if (GetPackage(edit: true) is not { } package) return false;
            string key = CreateImageContentKey(contentType, data);
            string partName;
            if (_packageImageParts != null && _packageImageParts.TryGetValue(key, out var reused) && package.PartExists(reused.Part)) {
                partName = reused.Part;
                width = reused.Width;
                height = reused.Height;
            } else {
                // The model numbers new media after the package-level parts in the folder, which never holds any, so it
                // asks for image1 and the SDK numbers on from there when that name is taken.
                string normalized = extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension;
                partName = UniquePartName(package, contentType, "/ppt/media/image1" + normalized);
                if (!PowerPointImageSize.TryRead(data, out width, out height)) width = height = 0;
                package.CreateBinaryPart(partName, null, data);
                string? declared = package.GetContentType(partName);
                if (declared == null) package.AddDefaultContentType(normalized.Substring(1), contentType);
                else if (!string.Equals(declared, contentType, StringComparison.OrdinalIgnoreCase)) package.SetOverride(partName, contentType);
                (_packageImageParts ??= new Dictionary<string, (string, int, int)>(StringComparer.Ordinal))[key] = (partName, width, height);
            }

            // The slide's relationship to the part when it has one; otherwise a new one with the SDK's kind of id.
            string relationships = OfficeOpcPartNames.GetRelationshipsPartName(slidePart);
            if (package.PartExists(relationships)) {
                OfficeXmlTree tree = package.GetXml(relationships);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                     node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    string? target = OfficeOpcPartNames.ResolveRelationshipTarget(relationships, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                    if (target != null && string.Equals(PartIdentity(target), PartIdentity(partName), StringComparison.OrdinalIgnoreCase)) {
                        relationshipId = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id") ?? string.Empty;
                        if (relationshipId.Length > 0) return true;
                    }
                }
            }
            relationshipId = package.AddRelationship(slidePart, ImageRelationshipType, OfficeOpcPartNames.GetRelativeTarget(slidePart, partName),
                id: "R" + Guid.NewGuid().ToString("N").Substring(0, 16));
            _pendingRelationships = null;
            return true;
        }
    }
}
