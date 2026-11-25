using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// A Word package held by OfficeIMO's own Open XML layer: the loaded bytes, the editable package over them, and the
/// main document part as an element tree, parsed on first use. Parts the engine does not model stay as they were read
/// and are written back untouched.
/// </summary>
internal sealed class WordPackageModel : IDisposable {
    internal const string DocumentMainContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    internal const string DigitalSignatureOriginRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
    private const string ExtendedPropertiesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "extended-properties";
    private const string MainDocumentPartName = "/word/document.xml";
    private const string CompatibilityMetadataPart = "/OfficeIMOCompatibility/SourceMetadata";
    private const string CompatibilityPayloadPart = "/OfficeIMOCompatibility/SourcePayload";
    private const string XmlSignatureContentType = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
    private const long MaximumApplicationPropertiesBytes = 4 * 1024 * 1024;

    private readonly MemoryStream _source;
    private OfficeXmlTree? _document;
    private int _body = OfficeXmlTree.NoNode;

    private WordPackageModel(byte[] sourceBytes, MemoryStream source, OfficeOpcEditablePackage package, string mainPartName) {
        SourceBytes = sourceBytes;
        _source = source;
        Package = package;
        MainPartName = mainPartName;
    }

    /// <summary>The bytes the package was loaded from.</summary>
    internal byte[] SourceBytes { get; }

    internal OfficeOpcEditablePackage Package { get; }

    internal string MainPartName { get; }

    internal OfficeXmlNamespaces.Table Namespaces => Package.Namespaces;

    /// <summary>The main document part, parsed on first use.</summary>
    internal OfficeXmlTree Document => _document ??= Package.GetXml(MainPartName);

    /// <summary>Whether the main document part has been parsed.</summary>
    internal bool IsDocumentParsed => _document != null;

    /// <summary>The <c>w:body</c> element of the main document part.</summary>
    internal int Body {
        get {
            if (_body == OfficeXmlTree.NoNode) {
                OfficeXmlTree document = Document;
                _body = document.FindChild(document.Root, OfficeXmlNamespaces.WordprocessingML, "body");
                if (_body == OfficeXmlTree.NoNode) throw new InvalidDataException("The main document part has no body.");
            }
            return _body;
        }
    }

    /// <summary>Whether anything was edited since the package was loaded.</summary>
    internal bool IsChanged { get; private set; }

    /// <summary>Whether the body has been brought to the state an SDK load leaves it in; see <see cref="WordOwnedEdits.PrepareBodyEdit"/>.</summary>
    internal bool BodyPrepared { get; set; }

    /// <summary>The number of <c>w:p</c> children of the body, counted on first use and kept current by owned edits.</summary>
    internal int BodyParagraphCount { get; set; } = -1;

    internal int GetBodyParagraphCount() {
        if (BodyParagraphCount < 0) {
            int count = 0;
            OfficeXmlTree document = Document;
            for (int child = document.FirstChildOf(Body); child != OfficeXmlTree.NoNode; child = document.NextSiblingOf(child)) {
                if (document.IsElement(child, OfficeXmlNamespaces.WordprocessingML, "p")) count++;
            }
            BodyParagraphCount = count;
        }
        return BodyParagraphCount;
    }

    /// <summary>Whether paragraph objects were handed out, which may be read after the document is disposed.</summary>
    internal bool ParagraphsHandedOut { get; set; }

    /// <summary>The bytes of the last save without the SDK, and the main part's version they were written from.</summary>
    internal byte[]? LastSavedBytes { get; set; }
    internal int LastSavedVersion { get; set; } = -1;

    /// <summary>The package as it stands, reusing the last save's bytes when nothing changed since.</summary>
    internal byte[] CurrentBytes() {
        if (!IsChanged) return SourceBytes;
        if (LastSavedBytes != null && IsDocumentParsed && Document.Version == LastSavedVersion) return LastSavedBytes;
        return ToBytes();
    }

    /// <summary>Records an edit, so the package is written from the model instead of copied.</summary>
    internal void MarkChanged() => IsChanged = true;

    /// <summary>
    /// Opens a plain, unsigned Transitional Word document: the usual main content type, a main part named as Word names
    /// it, no signature, and a <c>w:document</c> root in the Transitional namespace. Returns null for anything else, which
    /// the SDK then opens exactly as before.
    /// </summary>
    internal static WordPackageModel? TryOpen(byte[] bytes, long maximumCharactersInPart) {
        if (bytes.Length < 4) return null;
        var source = new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: true);
        OfficeOpcEditablePackage? editable = null;
        try {
            string? main;
            using (OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true)) {
                main = FindPlainMainPart(package, maximumCharactersInPart);
            }
            if (main == null) {
                source.Dispose();
                return null;
            }
            source.Position = 0;
            editable = OfficeOpcEditablePackage.Open(source, leaveOpen: true, poolTrees: true);
            var model = new WordPackageModel(bytes, source, editable, main);
            // The main part is parsed now, so a part the owned reader can't parse goes to the SDK at load. Strict markup
            // goes through the SDK too, which translates it.
            OfficeXmlTree document = model.Document;
            if (!document.IsElement(document.Root, OfficeXmlNamespaces.WordprocessingML, "document")
                || document.UriIdOf(document.Root) != OfficeXmlNamespaces.TransitionalUriId(OfficeXmlNamespaces.WordprocessingML)) {
                model.Dispose();
                return null;
            }
            _ = model.Body;
            return model;
        } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException
            or ArgumentException or FormatException or NotSupportedException or InvalidOperationException or OverflowException) {
            editable?.Dispose();
            source.Dispose();
            return null;
        }
    }

    private static string? FindPlainMainPart(OfficeOpcPackage package, long maximumCharactersInPart) {
        foreach (OfficeOpcRelationship relationship in package.GetRelationships(null)) {
            if (string.Equals(relationship.Type, DigitalSignatureOriginRelationshipType, StringComparison.Ordinal)) return null;
        }
        // Signature parts the package advertises, even without an origin relationship, make the save policy apply.
        if (package.DeclaresContentType(XmlSignatureContentType)) return null;
        // A package carrying a compatibility source keeps it through the SDK path, which restores it on save.
        if (package.PartExists(CompatibilityMetadataPart) && package.PartExists(CompatibilityPayloadPart)) return null;
        string? main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
        // The engine's save path names the main part by its usual URI, so other names keep going through the SDK.
        if (main == null || !string.Equals(main, MainDocumentPartName, StringComparison.Ordinal) || !package.PartExists(main)
            || !string.Equals(package.GetContentType(main), DocumentMainContentType, StringComparison.OrdinalIgnoreCase)) {
            return null;
        }
        // A UTF-8 part has no more characters than bytes, so one within the limit in bytes is within it in characters.
        if (maximumCharactersInPart > 0 && package.GetPartLength(main) > maximumCharactersInPart) return null;
        string? application = package.GetRelatedPartName(null, ExtendedPropertiesRelationshipType);
        if (application != null && package.PartExists(application)) {
            long length = package.GetPartLength(application);
            if (length > MaximumApplicationPropertiesBytes || maximumCharactersInPart > 0 && length > maximumCharactersInPart) return null;
            if (HasDigitalSignatureFlag(package, application)) return null;
        }
        return main;
    }

    private static bool HasDigitalSignatureFlag(OfficeOpcPackage package, string application) {
        OfficeXmlPullReader reader = package.OpenXml(application, out OfficePooledBytes content, processMarkupCompatibility: false);
        using (content) {
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 2
                    && reader.IsElement(OfficeXmlNamespaces.ExtendedProperties, "DigSig")) {
                    return true;
                }
            }
        }
        return false;
    }

    /// <summary>The package as it stands: the loaded bytes when nothing changed, otherwise the package written again.</summary>
    internal byte[] ToBytes() {
        if (!IsChanged) return SourceBytes;
        using var output = new MemoryStream(SourceBytes.Length + 4096);
        Package.Save(output);
        return output.ToArray();
    }

    /// <summary>Writes the package to <paramref name="destination"/>.</summary>
    internal void Save(Stream destination) => Package.Save(destination);

    public void Dispose() {
        Package.Dispose();
        _source.Dispose();
    }
}
