using System.IO.Packaging;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word;

internal static partial class WordPackageSnapshot {
    /// <summary>
    /// Copies physical OPC parts once, retaining shared relationships, opaque payloads and pending XML edits.
    /// Unlike SDK 3.5.1 Clone, this does not independently import each package-root relationship's subgraph.
    /// Only already-loaded roots are serialized; the source package is neither saved nor normalized.
    /// </summary>
#pragma warning disable OOXML0001
    internal static WordprocessingDocument Clone(WordprocessingDocument source, Stream destination) {
        Dictionary<Uri, OpenXmlPart> loadedParts = WordPackageParts.Enumerate(source, int.MaxValue)
            .Where(part => part.IsRootElementLoaded)
            .ToDictionary(part => part.Uri);
        string? strictMainId = source.StrictRelationshipFound && source.MainDocumentPart != null
            ? source.GetIdOfPart(source.MainDocumentPart) : null;
        IPackage package = source.GetPackage();
        using (Package copy = Package.Open(destination, FileMode.Create, FileAccess.ReadWrite)) {
            foreach (IPackagePart part in package.GetParts()) {
                // Relationship collections include pending changes; raw .rels streams may still be stale.
                if (PackUriHelper.IsRelationshipPartUri(part.Uri)) continue;
                PackagePart copiedPart = copy.CreatePart(part.Uri, part.ContentType, CompressionOption.Normal);
                using (Stream output = copiedPart.GetStream(FileMode.Create, FileAccess.Write)) {
                    if (loadedParts.TryGetValue(part.Uri, out OpenXmlPart? loaded) && loaded.RootElement != null) {
                        loaded.RootElement.Save(output);
                    } else {
                        using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
                        input.CopyTo(output);
                    }
                }
                foreach (IPackageRelationship relationship in part.Relationships) {
                    copiedPart.CreateRelationship(relationship.TargetUri, relationship.TargetMode,
                        relationship.RelationshipType, relationship.Id);
                }
            }
            foreach (IPackageRelationship relationship in package.Relationships) {
                // The SDK exposes Strict relationships as Transitional. Retain the main Strict marker
                // so it also translates untouched Strict XML when saving the copy, without loading source roots.
                string type = relationship.Id == strictMainId
                    ? "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
                    : relationship.RelationshipType;
                copy.CreateRelationship(relationship.TargetUri, relationship.TargetMode,
                    type, relationship.Id);
            }
        }
        destination.Position = 0;
        WordprocessingDocument clone = WordprocessingDocument.Open(destination, true);
        try {
            OfficeOpenXmlPackagePayload.CopyCoreProperties(source, clone);
            return clone;
        } catch {
            clone.Dispose();
            throw;
        }
    }
#pragma warning restore OOXML0001
}
