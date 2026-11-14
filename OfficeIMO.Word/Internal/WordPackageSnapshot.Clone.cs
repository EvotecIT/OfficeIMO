using System.IO.Packaging;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word;

internal static partial class WordPackageSnapshot {
    // WordPackageCompatibility normalizes the relationships of this part and of the package.
    private const string MainDocumentPartUri = "/word/document.xml";

    /// <summary>
    /// Copies physical OPC parts once, retaining shared relationships, opaque payloads and pending XML edits.
    /// Unlike SDK 3.5.1 Clone, this does not independently import each package-root relationship's subgraph.
    /// Only already-loaded roots are serialized into the copy; the source DOM is neither saved nor normalized.
    /// </summary>
#pragma warning disable OOXML0001
    internal static WordprocessingDocument Clone(WordprocessingDocument source, Stream destination) {
        WriteCopy(source, destination, copyCoreProperties: false);
        destination.Position = 0;
        WordprocessingDocument clone = WordprocessingDocument.Open(destination, true, CreateOpenSettings(source, autoSave: true));
        try {
            OfficeOpenXmlPackagePayload.CopyCoreProperties(source, clone);
            return clone;
        } catch {
            clone.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Writes the same package as <see cref="Clone"/> followed by disposing the clone, without reopening it, when the
    /// copy needs no Open XML processing: the source has no Strict main relationship to translate.
    /// </summary>
    /// <param name="source">Package to copy.</param>
    /// <param name="destination">Empty stream receiving the copy.</param>
    /// <param name="needsOpenOfficeNormalization">Whether a root or main-document relationship has an absolute target
    /// that <see cref="WordPackageCompatibility.NormalizeOpenOfficeRelationships(Stream)"/> would rewrite.</param>
    /// <returns><see langword="false"/> without writing when the caller must use <see cref="Clone"/>.</returns>
    internal static bool TryWriteFinalCopy(WordprocessingDocument source, Stream destination, out bool needsOpenOfficeNormalization) {
        needsOpenOfficeNormalization = false;
        if (source.StrictRelationshipFound) return false;
        needsOpenOfficeNormalization = WriteCopy(source, destination, copyCoreProperties: true);
        return true;
    }

    private static bool WriteCopy(WordprocessingDocument source, Stream destination, bool copyCoreProperties) {
        Dictionary<Uri, OpenXmlPart> loadedParts = WordPackageParts.Enumerate(source, int.MaxValue)
            .Where(part => part.IsRootElementLoaded)
            .ToDictionary(part => part.Uri);
        string? strictMainId = source.StrictRelationshipFound && source.MainDocumentPart != null
            ? source.GetIdOfPart(source.MainDocumentPart) : null;
        IPackage package = source.GetPackage();
        bool needsOpenOfficeNormalization = false;
        // Flush already-written streams and pending .rels before the Framework URI adapter enumerates them.
        // Saving the Open XML document here would also persist/normalize its loaded DOMs.
        if (package.FileOpenAccess == FileAccess.ReadWrite) package.Save();
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
                bool mainDocument = string.Equals(part.Uri.OriginalString, MainDocumentPartUri, StringComparison.OrdinalIgnoreCase);
                foreach (IPackageRelationship relationship in part.Relationships) {
                    needsOpenOfficeNormalization |= mainDocument && TargetStartsWith(relationship, "/word/");
                    copiedPart.CreateRelationship(relationship.TargetUri, relationship.TargetMode,
                        relationship.RelationshipType, relationship.Id);
                }
            }
            foreach (IPackageRelationship relationship in package.Relationships) {
                needsOpenOfficeNormalization |= TargetStartsWith(relationship, "/");
                // The SDK exposes Strict relationships as Transitional. Retain the main Strict marker
                // so it also translates untouched Strict XML when saving the copy, without loading source roots.
                string type = relationship.Id == strictMainId
                    ? "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
                    : relationship.RelationshipType;
                copy.CreateRelationship(relationship.TargetUri, relationship.TargetMode,
                    type, relationship.Id);
            }
            if (copyCoreProperties) CopyCoreProperties(source.PackageProperties, copy.PackageProperties);
        }
        return needsOpenOfficeNormalization;
    }

    /// <summary>Whether <see cref="WordPackageCompatibility.NormalizeOpenOfficeRelationships(Stream)"/> would rewrite a relationship of this package.</summary>
    internal static bool NeedsOpenOfficeNormalization(WordprocessingDocument document) {
        IPackage package = document.GetPackage();
        foreach (IPackageRelationship relationship in package.Relationships) {
            if (TargetStartsWith(relationship, "/")) return true;
        }
        var mainDocumentUri = new Uri(MainDocumentPartUri, UriKind.Relative);
        if (!package.PartExists(mainDocumentUri)) return false;
        foreach (IPackageRelationship relationship in package.GetPart(mainDocumentUri).Relationships) {
            if (TargetStartsWith(relationship, "/word/")) return true;
        }
        return false;
    }

    /// <summary>Whether <see cref="WordPackageCompatibility.NormalizeOpenOfficeRelationships(Stream)"/> would rewrite this target.</summary>
    private static bool TargetStartsWith(IPackageRelationship relationship, string prefix) =>
        relationship.TargetUri.OriginalString.StartsWith(prefix, StringComparison.Ordinal);

    /// <summary>Mirrors <see cref="OfficeOpenXmlPackagePayload.CopyCoreProperties"/> onto the package being written.</summary>
    private static void CopyCoreProperties(IPackageProperties src, PackageProperties dest) {
        dest.Category = src.Category;
        dest.ContentStatus = src.ContentStatus;
        dest.ContentType = src.ContentType;
        dest.Created = src.Created;
        dest.Creator = src.Creator;
        dest.Description = src.Description;
        dest.Identifier = src.Identifier;
        dest.Keywords = src.Keywords;
        dest.Language = src.Language;
        dest.LastModifiedBy = src.LastModifiedBy;
        dest.LastPrinted = src.LastPrinted;
        dest.Modified = src.Modified;
        dest.Revision = src.Revision;
        dest.Subject = src.Subject;
        dest.Title = src.Title;
        dest.Version = src.Version;
    }

    /// <summary>Retains parsing policy when opening a writable snapshot or reopening its saved bytes.</summary>
    internal static OpenSettings CreateOpenSettings(WordprocessingDocument source, bool autoSave) {
        MarkupCompatibilityProcessSettings compatibility = source.MarkupCompatibilityProcessSettings;
        return new OpenSettings {
            AutoSave = autoSave,
            MaxCharactersInPart = source.MaxCharactersInPart,
            MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                compatibility.ProcessMode, compatibility.TargetFileFormatVersions)
        };
    }

#pragma warning restore OOXML0001
}


