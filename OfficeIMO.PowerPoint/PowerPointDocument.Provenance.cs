using System.IO.Compression;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.Provenance;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    /// <summary>Validates and inspects encoded presentation bytes without accessing the filesystem.</summary>
    internal static OfficeProvenanceReport InspectProvenance(byte[] data, string fileName = "presentation.pptx", OfficeProvenanceOptions? options = null) =>
        OfficeProvenancePackageMutation.Inspect(data, fileName, options, ValidatePackage);

    /// <summary>Inspects C2PA and IPTC provenance in a saved Open XML presentation and its supported embedded images.</summary>
    internal static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenancePackageMutation.InspectFile(filePath, options, ValidatePackage);

    /// <summary>Removes selected provenance and atomically writes an Open XML presentation.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    /// <summary>Removes selected provenance from encoded Open XML presentation bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] presentationBytes,
        string fileName = "presentation.pptx",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(presentationBytes, fileName, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    private const string SignatureOriginContentType = "application/vnd.openxmlformats-package.digital-signature-origin";
    private const string SignatureRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
    private const string SignatureContentType = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
    private const string PackageRelationshipsPart = "/_rels/.rels";
    private const string ExtendedPropertiesNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    private static void ValidatePackage(byte[] data, string fileName, OfficeProvenanceOptions options) {
        OfficeProvenanceZip.ValidateForOwningPackageMutation(data, options);
        using var stream = new MemoryStream(data, writable: false);
        using OfficeOpcPackage package = OpenProvenancePackage(stream, options);
        string expectedExtension = PowerPointFormatCatalog.GetByExtension(fileName).Extension;
        string? main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
        string? contentType = main != null && package.PartExists(main) ? package.GetContentType(main) : null;
        string? actualExtension = PresentationExtension(contentType);
        if (actualExtension == null && PresentationSubtypes.ContentTypes.Any(subtype =>
                string.Equals(subtype.ContentType, contentType, StringComparison.OrdinalIgnoreCase))) {
            // The SDK falls back to Presentation when the known MIME differs only in spelling.
            actualExtension = ".pptx";
        }
        if (actualExtension == null) throw new InvalidDataException("The package is not a PowerPoint presentation.");

        if (!string.Equals(expectedExtension, actualExtension, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException($"The PowerPoint package subtype '{actualExtension}' does not match filename extension '{expectedExtension}'.");
        }
        if (options.RequireStandardOpenXmlDocument && actualExtension != ".pptx")
            throw new InvalidDataException("The memory-only workflow requires a PPTX presentation, not a macro-enabled presentation, template, or slideshow.");

        FindProvenancePart(package, OfficeOpcPackage.CorePropertiesRelationshipType, PowerPointPackageModel.CorePropertiesContentType);
        FindProvenancePart(package, ExtendedPropertiesRelationshipType, PowerPointPackageModel.ExtendedPropertiesContentType);
        FindProvenancePart(package, DigitalSignatureOriginRelationshipType, SignatureOriginContentType);
    }

    private static OfficeOpcPackage OpenProvenancePackage(Stream stream, OfficeProvenanceOptions limits) =>
        OfficeOpcPackage.Open(stream, maximumEntryLength: limits.MaxAssetBytes, maximumEntries: limits.MaxContainerEntries);

    private static class PresentationSubtypes {
        internal static readonly (string ContentType, string Extension)[] ContentTypes = [
            (PresentationMainContentType, ".pptx"),
            ("application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml", ".pptm"),
            ("application/vnd.ms-powerpoint.addin.macroEnabled.main+xml", ".ppam"),
            ("application/vnd.openxmlformats-officedocument.presentationml.template.main+xml", ".potx"),
            ("application/vnd.ms-powerpoint.template.macroEnabled.main+xml", ".potm"),
            ("application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml", ".ppsx"),
            ("application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml", ".ppsm")
        ];
    }

    private static string? PresentationExtension(string? contentType) {
        foreach ((string type, string extension) in PresentationSubtypes.ContentTypes) {
            if (string.Equals(type, contentType, StringComparison.Ordinal)) return extension;
        }
        return null;
    }

    private static OfficeOpcRelationship? FindProvenancePart(OfficeOpcPackage package, string relationshipType, string contentType) {
        OfficeOpcRelationship? first = null;
        foreach (OfficeOpcRelationship relationship in package.GetRelationships(null)) {
            if (relationship.IsExternal || relationship.Type != relationshipType) continue;
            if (relationship.TargetPartName is not { } target || !package.PartExists(target) ||
                !string.Equals(package.GetContentType(target), contentType, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException($"The PowerPoint relationship '{relationship.Id}' does not target a '{contentType}' part.");
            }
            first ??= relationship;
        }
        return first;
    }

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) =>
        OfficeProvenanceZip.HasNativePackageSignature(data, options) ||
        HasRelationshipOwnedApplicationSignatureMetadata(data, options.Limits);

    private static bool HasRelationshipOwnedApplicationSignatureMetadata(byte[] data, OfficeProvenanceOptions limits) {
        using var stream = new MemoryStream(data, writable: false);
        using OfficeOpcPackage package = OpenProvenancePackage(stream, limits);
        string? application = FindProvenancePart(package, ExtendedPropertiesRelationshipType, PowerPointPackageModel.ExtendedPropertiesContentType)?.TargetPartName;
        return HasApplicationSignatureMetadata(package, application, limits);
    }

    private static bool HasApplicationSignatureMetadata(OfficeOpcPackage package, string? application, OfficeProvenanceOptions limits) {
        if (application == null) return false;
        if (package.GetPartLength(application) > limits.MaxAssetBytes)
            throw new InvalidDataException("Open XML application metadata exceeds the configured asset limit.");

        using OfficePooledBytes bytes = package.ReadPart(application);
        if (bytes.Length == 0) return false;
        using var stream = new MemoryStream(bytes.Array, 0, bytes.Length, writable: false);
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(stream, limits, "Open XML application metadata");
        stream.Position = 0;

        // The SDK reads a direct typed property. Same-named elements in producer extensions or alternate
        // content cannot supply a signature or be removed with the property.
        using var reader = XmlReader.Create(stream, OfficeProvenanceXml.CreateReaderSettings(limits));
        if (reader.MoveToContent() != XmlNodeType.Element ||
            reader.LocalName != "Properties" || reader.NamespaceURI != ExtendedPropertiesNamespace) {
            throw new InvalidDataException("Cannot load the root element from the part. The part contains invalid data.");
        }
        while (reader.Read()) {
            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1 &&
                reader.LocalName == "DigSig" && reader.NamespaceURI == ExtendedPropertiesNamespace) return true;
        }
        return false;
    }

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) {
        OfficeProvenanceOptions limits = options.Limits;
        using var source = new MemoryStream(data, writable: false);
        using OfficeOpcPackage package = OpenProvenancePackage(source, limits);
        OfficeOpcRelationship? origin = FindProvenancePart(package, DigitalSignatureOriginRelationshipType, SignatureOriginContentType);
        string? application = FindProvenancePart(package, ExtendedPropertiesRelationshipType, PowerPointPackageModel.ExtendedPropertiesContentType)?.TargetPartName;
        bool hasApplicationMetadata = HasApplicationSignatureMetadata(package, application, limits);
        if (origin == null && !hasApplicationMetadata) return new OfficeProvenanceSignatureStripResult(data, hadSignatures: false);

        HashSet<string> removed = GetExclusiveSignatureParts(package, origin, limits.CancellationToken);
        string? applicationIdentity = hasApplicationMetadata ? OfficeOpcPackage.GetPartIdentity(application!) : null;
        using var output = new OfficeProvenanceBoundedMemoryStream(options.EffectiveMaxIntermediateBytes, data.Length);
        output.Write(data, 0, data.Length);
        output.Position = 0;
        // Updating the existing ZIP keeps the comments, extra fields and attributes of retained entries.
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            var writer = new OfficeXmlUtf8Writer();
            foreach (ZipArchiveEntry entry in archive.Entries.ToArray()) {
                limits.CancellationToken.ThrowIfCancellationRequested();
                string part = OfficeOpcPackage.GetPartIdentity("/" + entry.FullName);
                if (removed.Contains(part)) {
                    entry.Delete();
                } else if (origin != null && OfficeAsciiIgnoreCaseComparer.Instance.Equals(part, PackageRelationshipsPart)) {
                    RewriteSignatureMetadata(package, entry, PackageRelationshipsPart, writer, tree => RemoveSignatureRelationship(tree, origin.Id));
                } else if (origin != null && OfficeAsciiIgnoreCaseComparer.Instance.Equals(part, OfficeOpcEditablePackage.ContentTypesPartName)) {
                    RewriteSignatureMetadata(package, entry, OfficeOpcEditablePackage.ContentTypesPartName, writer, tree => RemoveSignatureOverrides(tree, removed));
                } else if (applicationIdentity != null && OfficeAsciiIgnoreCaseComparer.Instance.Equals(part, applicationIdentity)) {
                    RewriteSignatureMetadata(package, entry, application!, writer, RemoveApplicationSignatureMetadata);
                }
            }
        }
        return new OfficeProvenanceSignatureStripResult(output.ToArray(), hadSignatures: true);
    }

    private static void RewriteSignatureMetadata(OfficeOpcPackage package, ZipArchiveEntry entry, string partName,
        OfficeXmlUtf8Writer writer, Action<OfficeXmlTree> edit) {
        using OfficePooledBytes bytes = package.ReadPart(partName);
        OfficeXmlTree tree = OfficeXmlTree.Parse(bytes.Array, bytes.Length, package.Namespaces);
        edit(tree);
        if (!tree.IsChanged) return;
        using Stream stream = entry.Open();
        stream.SetLength(0);
        writer.Target = stream;
        tree.WriteTo(writer);
        writer.Flush();
        writer.Target = null;
    }

    private static void RemoveApplicationSignatureMetadata(OfficeXmlTree tree) {
        const int ns = OfficeXmlNamespaces.ExtendedProperties;
        for (int signature = tree.FindChild(tree.Root, ns, "DigSig"); signature != OfficeXmlTree.NoNode;) {
            int next = tree.FindNextSibling(signature, ns, "DigSig");
            if (tree.UriIdOf(signature) == OfficeXmlNamespaces.TransitionalUriId(ns)) tree.Remove(signature);
            signature = next;
        }
    }

    private static void RemoveSignatureRelationship(OfficeXmlTree tree, string relationshipId) {
        const int ns = OfficeXmlNamespaces.PackageRelationships;
        for (int element = tree.FindChild(tree.Root, ns, "Relationship"); element != OfficeXmlTree.NoNode;
             element = tree.FindNextSibling(element, ns, "Relationship")) {
            if (tree.GetAttribute(element, OfficeXmlNamespaces.None, "Id") != relationshipId) continue;
            tree.Remove(element);
            return;
        }
    }

    private static void RemoveSignatureOverrides(OfficeXmlTree tree, HashSet<string> removed) {
        const int ns = OfficeXmlNamespaces.ContentTypes;
        for (int element = tree.FindChild(tree.Root, ns, "Override"); element != OfficeXmlTree.NoNode;) {
            int next = tree.FindNextSibling(element, ns, "Override");
            if (tree.GetAttribute(element, OfficeXmlNamespaces.None, "PartName") is { } path &&
                removed.Contains(OfficeOpcPackage.GetPartIdentity(path))) tree.Remove(element);
            element = next;
        }
    }

    private static HashSet<string> GetExclusiveSignatureParts(OfficeOpcPackage package, OfficeOpcRelationship? origin,
        CancellationToken cancellationToken) {
        var types = new Dictionary<string, string>(OfficeAsciiIgnoreCaseComparer.Instance);
        void CheckType(string target, string relationshipType) {
            if (types.TryGetValue(target, out string? prior) && prior != relationshipType)
                throw new InvalidDataException("A signature part is shared by relationships with different types.");
        }

        if (origin == null) {
            VisitReachablePackageParts(package, null, null, cancellationToken);
            return new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
        }

        string start = origin.TargetPartName!;
        types.Add(OfficeOpcPackage.GetPartIdentity(start), origin.Type);
        VisitReachablePackageParts(package, start, null, cancellationToken, (source, relationship, target) => {
            CheckType(target, relationship.Type);
            types[target] = relationship.Type;
            if (source == start && relationship.Type == SignatureRelationshipType &&
                !string.Equals(package.GetContentType(relationship.TargetPartName!), SignatureContentType, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException($"The PowerPoint relationship '{relationship.Id}' does not target a '{SignatureContentType}' part.");
            }
        });

        HashSet<string> kept = VisitReachablePackageParts(package, null, origin.Id, cancellationToken,
            (_, relationship, target) => CheckType(target, relationship.Type));
        var removed = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
        foreach (string signature in types.Keys) {
            if (kept.Contains(signature)) continue;
            removed.Add(signature);
            removed.Add(OfficeOpcPartNames.GetRelationshipsPartName(signature));
        }
        return removed;
    }

    // Check each incoming relationship, while opening each part's relationships once per walk.
    private static HashSet<string> VisitReachablePackageParts(OfficeOpcPackage package, string? start, string? skipRootId,
        CancellationToken cancellationToken, Action<string?, OfficeOpcRelationship, string>? visit = null) {
        var reached = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
        if (start != null) reached.Add(OfficeOpcPackage.GetPartIdentity(start));
        var pending = new Stack<string?>();
        pending.Push(start);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            string? source = pending.Pop();
            foreach (OfficeOpcRelationship relationship in package.GetRelationships(source)) {
                if (relationship.IsExternal || relationship.Type == PowerPointPartRef.HyperlinkRelationshipType ||
                    source == null && relationship.Id == skipRootId) continue;
                if (relationship.TargetPartName is not { } target || !package.PartExists(target)) {
                    throw new InvalidDataException($"The PowerPoint relationship '{relationship.Id}' has no target part.");
                }
                string identity = OfficeOpcPackage.GetPartIdentity(target);
                visit?.Invoke(source, relationship, identity);
                if (reached.Add(identity)) pending.Push(target);
            }
        }
        return reached;
    }
}
