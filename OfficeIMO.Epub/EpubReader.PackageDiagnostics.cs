namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    private const string IdpfFontObfuscationAlgorithm = "http://www.idpf.org/2008/embedding";
    private const string AdobeFontObfuscationAlgorithm = "http://ns.adobe.com/pdf/enc#RC";

    private static IReadOnlyList<EpubEncryptionInfo> ReadEncryption(
        IReadOnlyDictionary<string, ZipArchiveEntry> entryIndex,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        const string encryptionPath = "META-INF/encryption.xml";
        if (!entryIndex.TryGetValue(encryptionPath, out ZipArchiveEntry? entry)) {
            return Array.Empty<EpubEncryptionInfo>();
        }
        if (entry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.encryption.metadata-size-limit",
                $"EPUB encryption.xml exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                encryptionPath);
            return Array.Empty<EpubEncryptionInfo>();
        }

        string content = ReadEntryText(entry, options.MaxPackageMetadataBytes);
        if (!TryParseXml(content, out XDocument? document) || document == null) {
            diagnostics.Warning(
                "epub.encryption.invalid-xml",
                "EPUB encryption.xml could not be parsed as XML.",
                encryptionPath);
            return Array.Empty<EpubEncryptionInfo>();
        }

        var results = new List<EpubEncryptionInfo>();
        var seenPaths = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement encryptedData in document.Descendants().Where(element => IsName(element, "EncryptedData"))) {
            XElement? method = encryptedData.Descendants().FirstOrDefault(element => IsName(element, "EncryptionMethod"));
            XElement? reference = encryptedData.Descendants().FirstOrDefault(element => IsName(element, "CipherReference"));
            string? algorithm = method == null ? null : NullIfWhiteSpace(GetAttribute(method, "Algorithm"));
            string uri = reference == null ? string.Empty : GetAttribute(reference, "URI");
            string resourcePath = ResolveContainerRootPath(uri);
            if (resourcePath.Length == 0) {
                diagnostics.Warning(
                    "epub.encryption.resource-path-invalid",
                    $"Ignored encryption declaration with invalid resource URI '{uri}'.",
                    encryptionPath);
                continue;
            }
            if (!seenPaths.Add(resourcePath)) {
                diagnostics.Warning(
                    "epub.encryption.duplicate-resource",
                    $"Ignored duplicate encryption declaration for '{resourcePath}'.",
                    resourcePath);
                continue;
            }

            EpubEncryptionKind kind = ClassifyEncryption(algorithm);
            var info = new EpubEncryptionInfo {
                Path = resourcePath,
                Algorithm = algorithm,
                Kind = kind
            };
            results.Add(info);

            if (!entryIndex.ContainsKey(resourcePath)) {
                diagnostics.Warning(
                    "epub.encryption.resource-missing",
                    $"Encryption declaration references missing resource '{resourcePath}'.",
                    resourcePath);
            } else if (info.RequiresDecryption) {
                diagnostics.Warning(
                    "epub.encryption.unsupported",
                    $"Resource '{resourcePath}' uses unsupported encryption algorithm '{algorithm ?? "(missing)"}'.",
                    resourcePath);
            } else {
                diagnostics.Info(
                    "epub.encryption.font-obfuscation",
                    $"Resource '{resourcePath}' uses recognized EPUB font obfuscation.",
                    resourcePath);
            }
        }
        return results;
    }

    private static EpubEncryptionKind ClassifyEncryption(string? algorithm) {
        if (string.Equals(algorithm, IdpfFontObfuscationAlgorithm, StringComparison.Ordinal)) {
            return EpubEncryptionKind.IdpfFontObfuscation;
        }
        if (string.Equals(algorithm, AdobeFontObfuscationAlgorithm, StringComparison.Ordinal)) {
            return EpubEncryptionKind.AdobeFontObfuscation;
        }
        return string.IsNullOrWhiteSpace(algorithm)
            ? EpubEncryptionKind.Unknown
            : EpubEncryptionKind.Encryption;
    }

    private static string ResolveContainerRootPath(string value) {
        string reference = RemoveFragmentAndQuery(value);
        if (reference.Length == 0 || reference.StartsWith("//", StringComparison.Ordinal)) return string.Empty;
        if (Uri.TryCreate(reference, UriKind.Absolute, out _)) return string.Empty;
        return CollapsePathSegments(NormalizePath(reference).TrimStart('/'));
    }

    private static EpubRenditionLayout? ReadPackageRenditionLayout(
        XElement metadata,
        EpubDiagnosticCollector diagnostics,
        string opfPath) {
        XElement[] declarations = metadata.Elements()
            .Where(element => IsName(element, "meta") &&
                string.Equals(GetAttribute(element, "property"), "rendition:layout", StringComparison.Ordinal))
            .ToArray();
        if (declarations.Length > 1) {
            diagnostics.Warning(
                "epub.layout.multiple-declarations",
                "EPUB package declares rendition:layout more than once. The first valid declaration is used.",
                opfPath);
        }

        foreach (XElement declaration in declarations) {
            if (TryParseRenditionLayout(declaration.Value, out EpubRenditionLayout layout)) return layout;
            diagnostics.Warning(
                "epub.layout.invalid",
                $"EPUB package declares unsupported rendition:layout value '{NormalizeWhitespace(declaration.Value)}'.",
                opfPath);
        }

        XElement? legacy = metadata.Elements().FirstOrDefault(element =>
            IsName(element, "meta") &&
            string.Equals(GetAttribute(element, "name"), "fixed-layout", StringComparison.OrdinalIgnoreCase));
        if (legacy != null) {
            string content = GetAttribute(legacy, "content");
            if (string.Equals(content, "true", StringComparison.OrdinalIgnoreCase)) return EpubRenditionLayout.PrePaginated;
            if (string.Equals(content, "false", StringComparison.OrdinalIgnoreCase)) return EpubRenditionLayout.Reflowable;
        }
        return null;
    }

    private static EpubRenditionLayout? ResolveRenditionLayout(
        EpubRenditionLayout? packageLayout,
        string? spineProperties) {
        if (ContainsSpaceSeparatedToken(spineProperties, "rendition:layout-pre-paginated")) {
            return EpubRenditionLayout.PrePaginated;
        }
        if (ContainsSpaceSeparatedToken(spineProperties, "rendition:layout-reflowable")) {
            return EpubRenditionLayout.Reflowable;
        }
        return packageLayout;
    }

    private static bool TryParseRenditionLayout(string? value, out EpubRenditionLayout layout) {
        string normalized = NormalizeWhitespace(value ?? string.Empty);
        if (string.Equals(normalized, "pre-paginated", StringComparison.OrdinalIgnoreCase)) {
            layout = EpubRenditionLayout.PrePaginated;
            return true;
        }
        if (string.Equals(normalized, "reflowable", StringComparison.OrdinalIgnoreCase)) {
            layout = EpubRenditionLayout.Reflowable;
            return true;
        }
        layout = default;
        return false;
    }

    private static EpubReadException CreateFatalReadException(
        string code,
        string message,
        string? path = null,
        Exception? innerException = null) {
        return new EpubReadException(
            message,
            new[] {
                new EpubDiagnostic {
                    Code = code,
                    Severity = EpubDiagnosticSeverity.Error,
                    Message = message,
                    Path = path
                }
            },
            innerException);
    }

    private sealed class EpubDiagnosticCollector {
        private readonly List<EpubDiagnostic> _items = new List<EpubDiagnostic>();

        public IReadOnlyList<EpubDiagnostic> Items => _items.ToArray();

        public IReadOnlyList<string> WarningMessages => _items
            .Where(static item => item.Severity != EpubDiagnosticSeverity.Info)
            .Select(static item => item.Message)
            .ToArray();

        public void Info(string code, string message, string? path = null) =>
            Add(code, EpubDiagnosticSeverity.Info, message, path);

        public void Warning(string code, string message, string? path = null) =>
            Add(code, EpubDiagnosticSeverity.Warning, message, path);

        private void Add(string code, EpubDiagnosticSeverity severity, string message, string? path) {
            _items.Add(new EpubDiagnostic {
                Code = code,
                Severity = severity,
                Message = message,
                Path = path
            });
        }
    }
}
