using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private string? _xmlSourcePath;
    private WeakReference<Stream>? _xmlSourceStream;
    private IReadOnlyList<OfficeDiagnostic> _xmlSourceDiagnostics = Array.Empty<OfficeDiagnostic>();

    private WordDocument RememberXmlStream(Stream source) {
        _xmlSourceStream = new WeakReference<Stream>(source);
        if (source is FileStream file) _xmlSourcePath = file.Name;
        return this;
    }

    private void EnsureSeparateXmlDestination(string path) {
        if (_xmlSourcePath != null && OfficePathIdentity.AreEquivalent(_xmlSourcePath, path)) {
            throw new InvalidOperationException("Save Word XML imports to a destination separate from the source.");
        }
    }

    private void EnsureSeparateXmlDestination(Stream destination) {
        if (_xmlSourceStream?.TryGetTarget(out Stream? source) == true && ReferenceEquals(source, destination)) {
            throw new InvalidOperationException("Save Word XML imports to a destination separate from the source.");
        }
        if (_xmlSourcePath != null && destination is FileStream file) EnsureSeparateXmlDestination(file.Name);
    }

    /// <summary>Exports a Word Open XML package as a standalone Flat OPC XML document.</summary>
    /// <remarks>Does not change the associated destination. Signed packages are blocked by default.
    /// To export a password-protected source as clear XML, explicitly set the save password to null.</remarks>
    public byte[] ExportFlatOpc(WordSaveOptions? options = null) {
        if (ResolveSavePassword(options) != null) {
            throw new NotSupportedException("Flat OPC cannot carry password protection.");
        }
        byte[] bytes = ToWordBytes(WordFileFormat.Docx, options);
        OfficePackageSecurityInspector.Validate(bytes, new OfficePackageSecurityOptions {
            MaxPackageBytes = WordXmlInterchange.MaximumXmlBytes,
            MaxTotalUncompressedBytes = WordXmlInterchange.MaximumXmlBytes,
            MaxPartUncompressedBytes = WordXmlInterchange.MaximumXmlBytes
        });
        using var package = new MemoryStream(bytes, writable: false);
        using var word = WordprocessingDocument.Open(package, false);
        using var xml = new OfficeBoundedMemoryStream(WordXmlInterchange.MaximumXmlBytes);
        using (XmlWriter writer = XmlWriter.Create(xml, new XmlWriterSettings { Encoding = new UTF8Encoding(false) })) {
            XDocument flat = word.ToFlatOpcDocument();
            if (!flat.Nodes().OfType<XProcessingInstruction>().Any(instruction => instruction.Target == "mso-application"))
                flat.AddFirst(new XProcessingInstruction("mso-application", "progid=\"Word.Document\""));
            flat.Save(writer);
        }
        return xml.ToArray();
    }

    private OfficeFormatDescriptor GetXmlSourceFormatDescriptor() {
        if (SourceFormat == WordFileFormat.WordMl2003) {
            return new OfficeFormatDescriptor("Word.WordMl2003", ".xml", OfficeDocumentFamily.Word,
                OfficeDocumentKind.Document, OfficeFormatGeneration.Legacy, OfficeFormatEncoding.StandaloneXml, macroEnabled: true);
        }
        OfficeFormatDescriptor package = WordFormatCatalog.GetDescriptor(_wordprocessingDocument.DocumentType);
        string id = _wordprocessingDocument.DocumentType switch {
            WordprocessingDocumentType.MacroEnabledDocument => "Word.FlatOpcMacroEnabled",
            WordprocessingDocumentType.Template => "Word.FlatOpcTemplate",
            WordprocessingDocumentType.MacroEnabledTemplate => "Word.FlatOpcMacroEnabledTemplate",
            _ => "Word.FlatOpc"
        };
        return new OfficeFormatDescriptor(id, ".xml", package.Family, package.DocumentKind, package.Generation, OfficeFormatEncoding.StandaloneXml, package.IsMacroEnabled);
    }

    private static WordDocument? TryLoadWordXml(byte[] bytes, string? sourcePath, WordLoadOptions options, CancellationToken cancellationToken = default) {
        XDocument? xml = WordXmlInterchange.TryRead(bytes, options.MaxInputBytes, cancellationToken);
        if (xml == null) return null;
        if (options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose) {
            throw new NotSupportedException("Word XML input is detached. Save explicitly to a Word package or export Flat OPC.");
        }

        OfficePackageSecurityOptions policy = options.PackageSecurity?.Snapshot() ?? OfficePackageSecurityOptions.SecureDefaults;
        const long maximumPackageBytes = WordXmlInterchange.MaximumXmlBytes;
        policy.MaxPackageBytes = maximumPackageBytes;
        using var package = new OfficeBoundedMemoryStream(maximumPackageBytes);
        bool isFlatOpc = xml.Root!.Name == WordXmlInterchange.FlatOpcPackage;
        IReadOnlyList<OfficeDiagnostic> diagnostics = Array.Empty<OfficeDiagnostic>();
        if (isFlatOpc) {
            WordXmlInterchange.ValidateFlatOpc(xml.Root!, policy, cancellationToken);
            WordprocessingDocument flat;
            try {
                flat = WordprocessingDocument.FromFlatOpcDocument(xml, package, true);
            } catch (FormatException error) {
                throw new InvalidDataException("Flat OPC contains an invalid base64 payload.", error);
            }
            using (flat) {
                if (flat.MainDocumentPart == null) {
                    throw new InvalidDataException("The Flat OPC source does not have a Word main document part.");
                }
            }
        } else {
            diagnostics = WordMl2003Projection.WritePackage(xml.Root!, package, policy, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();

        package.Position = 0;
        WordLoadOptions packageOptions = options.Snapshot();
        packageOptions.PackageSecurity = policy;
        packageOptions.MaxInputBytes = maximumPackageBytes;
        WordDocument document = LoadFromStream(package, packageOptions, captureSourcePackageSecurity: true);
        document.OriginalStream = null!;
        document.SourceFormat = isFlatOpc ? WordFileFormat.FlatOpc : WordFileFormat.WordMl2003;
        document._xmlSourcePath = sourcePath;
        document._xmlSourceDiagnostics = diagnostics;
        if (document.SourcePackageSecurityReport != null) document.SourcePackageSecurityScope = "reconstructed-open-xml-package";
        if (options.FailOnLoss && diagnostics.Count > 0) {
            var report = document.CreateLoadReport(sourcePath ?? "Word XML stream", TimeSpan.Zero);
            document.Dispose();
            report.ThrowIfLoss();
        }
        return document;
    }
}

internal static class WordXmlInterchange {
    internal const long MaximumXmlBytes = 64L * 1024 * 1024;
    internal const int MaximumXmlElements = 250_000;
    private static readonly XNamespace FlatOpc = "http://schemas.microsoft.com/office/2006/xmlPackage";
    internal static readonly XNamespace WordMl = "http://schemas.microsoft.com/office/word/2003/wordml";
    internal static readonly XName FlatOpcPackage = FlatOpc + "package";
    private static readonly XName WordMlDocument = WordMl + "wordDocument";

    internal static XDocument? TryRead(byte[] bytes, long maximumInputBytes, CancellationToken token) {
        if (WordDocumentLoadRouting.HasZipSignature(bytes) || WordDocumentLoadRouting.HasOleCompoundSignature(bytes)) {
            return null;
        }
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = Math.Min(maximumInputBytes, MaximumXmlBytes)
        };
        using var source = new MemoryStream(bytes, writable: false);
        using var reader = new OfficeXmlLimitingReader(XmlReader.Create(source, settings), "Word XML",
            maxDepth: 128, maxElements: MaximumXmlElements, maxAttributes: 1_000_000, cancellationToken: token);
        try {
            if (reader.MoveToContent() != XmlNodeType.Element) return null;
        } catch (XmlException) {
            return null;
        }
        XName name = XName.Get(reader.LocalName, reader.NamespaceURI);
        if (name != FlatOpcPackage && name != WordMlDocument) return null;
        if (bytes.LongLength > MaximumXmlBytes) {
            throw new InvalidDataException("Word XML exceeds the supported input size.");
        }
        try {
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        } catch (XmlException error) {
            throw new InvalidDataException("Word XML contains invalid XML.", error);
        }
    }

    internal static void ValidateFlatOpc(XElement root, OfficePackageSecurityOptions policy, CancellationToken token) {
        var parts = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        long aggregateLimit = Math.Min(MaximumXmlBytes, policy.MaxTotalUncompressedBytes);
        long aggregate = 0;
        foreach (XElement part in root.Elements()) {
            token.ThrowIfCancellationRequested();
            if (part.Name != FlatOpc + "part") throw new InvalidDataException("Unexpected element in a Flat OPC package.");
            if (parts.Count >= policy.MaxPartCount) {
                throw new OfficePackageSecurityException(OfficePackageSecurityRule.PartCount,
                    "Flat OPC exceeds the package part limit.", parts.Count + 1, policy.MaxPartCount);
            }
            string name = (string?)part.Attribute(FlatOpc + "name") ?? string.Empty;
            string contentType = (string?)part.Attribute(FlatOpc + "contentType") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(contentType) || !Uri.TryCreate(name, UriKind.Relative, out Uri? uri) ||
                !name.StartsWith("/", StringComparison.Ordinal) || name.IndexOfAny(new[] { '\\', '#', '?' }) >= 0) {
                throw new InvalidDataException("Flat OPC contains an invalid part name or content type.");
            }
            Uri partUri;
            try {
                partUri = PackUriHelper.CreatePartUri(uri);
            } catch (ArgumentException error) {
                throw new InvalidDataException("Flat OPC contains an invalid package part name.", error);
            }
            if (!string.Equals(name, partUri.OriginalString, StringComparison.Ordinal) ||
                parts.ContainsKey(partUri.OriginalString)) {
                throw new InvalidDataException("Flat OPC contains an invalid or duplicate package part name.");
            }
            if (name.StartsWith("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase) ||
                contentType.IndexOf("digital-signature", StringComparison.OrdinalIgnoreCase) >= 0) {
                throw new NotSupportedException("Flat OPC import does not preserve package-level digital signatures.");
            }
            XElement? data = part.Elements().FirstOrDefault();
            if (data == null || data.ElementsAfterSelf().Any()) throw new InvalidDataException("Flat OPC part must have one payload.");
            long length;
            if (data.Name == FlatOpc + "xmlData") {
                XElement? payload = data.Elements().FirstOrDefault();
                if (payload == null || payload.ElementsAfterSelf().Any()) {
                    throw new InvalidDataException("Flat OPC XML part must have one document element.");
                }
                AddInheritedNamespaces(payload, aggregate, aggregateLimit);
                string text = payload.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
                if (text.Length > policy.MaxXmlCharactersInPart) {
                    throw new InvalidDataException("Flat OPC part exceeds the package XML character limit.");
                }
                length = Encoding.UTF8.GetByteCount(text);
            } else if (data.Name == FlatOpc + "binaryData" && !data.HasElements) {
                length = EstimateBinaryBytes(data.Value);
            } else {
                throw new InvalidDataException("Flat OPC contains an unsupported part payload.");
            }
            if (length > policy.MaxPartUncompressedBytes) {
                throw new OfficePackageSecurityException(OfficePackageSecurityRule.PartSize,
                    "Flat OPC part exceeds the decoded size limit.", length, policy.MaxPartUncompressedBytes);
            }
            if (length > aggregateLimit - aggregate) {
                throw new OfficePackageSecurityException(OfficePackageSecurityRule.TotalUncompressedSize,
                    "Flat OPC exceeds the aggregate decoded size limit.", aggregate + length, aggregateLimit);
            }
            aggregate += length;
            parts.Add(partUri.OriginalString, part);
        }
        EnsureWordMainPart(parts);
    }

    private static void AddInheritedNamespaces(XElement payload, long aggregate, long aggregateLimit) {
        var prefixes = new HashSet<XName>();
        long addedBytes = 0;
        foreach (XElement scope in payload.AncestorsAndSelf()) {
            foreach (XAttribute declaration in scope.Attributes()) {
                if (!declaration.IsNamespaceDeclaration || !prefixes.Add(declaration.Name) || scope == payload) continue;
                addedBytes += Encoding.UTF8.GetByteCount(declaration.Value) + declaration.Name.LocalName.Length + 10L;
                if (addedBytes > aggregateLimit - aggregate) {
                    throw new OfficePackageSecurityException(OfficePackageSecurityRule.TotalUncompressedSize,
                        "Inherited Flat OPC namespaces exceed the decoded package size limit.", aggregate + addedBytes, aggregateLimit);
                }
                payload.Add(new XAttribute(declaration));
            }
        }
    }

    private static void EnsureWordMainPart(Dictionary<string, XElement> parts) {
        XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement? rootRelationships = parts.TryGetValue("/_rels/.rels", out XElement? part)
            ? part.Element(FlatOpc + "xmlData")?.Element(relationships + "Relationships")
            : null;
        XElement[] rootTargets = rootRelationships?.Elements(relationships + "Relationship")
            .Where(element => (string?)element.Attribute("Type") is
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" or
                "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument").ToArray()
            ?? Array.Empty<XElement>();
        if (rootTargets.Length != 1 || (string?)rootTargets[0].Attribute("TargetMode") is not (null or "Internal")) {
            throw new InvalidDataException("Flat OPC must identify one internal Word main document part.");
        }

        string target = (string?)rootTargets[0].Attribute("Target") ?? string.Empty;
        if (!Uri.TryCreate(target, UriKind.Relative, out Uri? relativeTarget) ||
            target.StartsWith("//", StringComparison.Ordinal) || target.IndexOfAny(new[] { '\\', '#', '?' }) >= 0) {
            throw new InvalidDataException("Flat OPC contains an invalid Word main part target.");
        }
        Uri targetPart;
        try {
            targetPart = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relativeTarget);
        } catch (ArgumentException error) {
            throw new InvalidDataException("Flat OPC contains an invalid Word main part target.", error);
        }
        if (!parts.TryGetValue(targetPart.OriginalString, out XElement? main) ||
            (string?)main.Attribute(FlatOpc + "contentType") is not
            ("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" or
             "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml" or
             "application/vnd.ms-word.document.macroEnabled.main+xml" or
             "application/vnd.ms-word.template.macroEnabledTemplate.main+xml")) {
            throw new InvalidDataException("The Flat OPC source does not have a Word main document part.");
        }
    }

    private static long EstimateBinaryBytes(string value) {
        long symbols = 0;
        int padding = 0;
        foreach (char c in value) {
            if (c is ' ' or '\t' or '\r' or '\n') continue;
            symbols++;
            if (c == '=') padding++;
        }
        if (symbols % 4 != 0 || padding > 2 || padding > symbols) {
            throw new InvalidDataException("Flat OPC contains an invalid base64 payload.");
        }
        return (symbols / 4 * 3) - padding;
    }
}
