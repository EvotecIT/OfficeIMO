using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static void ProjectLegacyVbaProject(
            LegacyPptImportParts parts,
            LegacyPpt.LegacyPptDocument legacy) {
            if (legacy.VbaProject == null) return;
            parts.AddBinaryPart(parts.PresentationPart, LegacyVbaProjectRelationshipType, "/ppt/vbaProject.bin",
                LegacyVbaProjectContentType, legacy.VbaProject.GetBytes(), numbered: false, generatedId: true, out _);
        }

        /// <summary>
        /// Gives a projected package with a VBA project the macro-enabled presentation content type, as a macro-enabled
        /// presentation declares its main part.
        /// </summary>
        internal static byte[] ConvertProjectedVbaPackageToMacroEnabled(
            byte[] packageBytes, PowerPointLoadOptions loadOptions) {
            ValidatePackageSecurity(packageBytes, loadOptions);
            using var input = new MemoryStream(packageBytes, writable: false);
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
            string? main = package.GetRelationshipTarget(null, FindPackageRelationshipId(package, LegacyOfficeDocumentRelationshipType)
                ?? throw new InvalidDataException("The projected presentation has no main part."))
                ?? throw new InvalidDataException("The projected presentation has no main part.");
            if (string.Equals(package.GetContentType(main), LegacyMacroPresentationMainContentType, StringComparison.OrdinalIgnoreCase)) {
                return packageBytes;
            }
            package.SetOverride(main, LegacyMacroPresentationMainContentType);
            using var output = new MemoryStream();
            package.Save(output);
            return output.ToArray();
        }

        private static string? FindPackageRelationshipId(OfficeOpcEditablePackage package, string type) {
            OfficeXmlTree tree = package.GetRelationships(null);
            for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(OfficeXmlNamespaces.NormalizeRelationshipType(tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty),
                        type, StringComparison.Ordinal)) {
                    return tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                }
            }
            return null;
        }
    }
}
