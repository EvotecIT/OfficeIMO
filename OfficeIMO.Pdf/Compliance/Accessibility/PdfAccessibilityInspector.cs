using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Checks one exact PDF against the PDF/UA-1 conditions in <see cref="PdfAccessibilityConditions"/>.
/// Every check reads the file as it is; nothing assumes OfficeIMO produced it.
/// </summary>
internal sealed partial class PdfAccessibilityInspector {
    private const int MaxFindingsPerCondition = 100;
    private const int MaxStructureDepth = 512;
    private static readonly Regex LanguagePattern = new("^[a-zA-Z]{1,8}(-[a-zA-Z0-9]{1,8})*$", RegexOptions.CultureInvariant);

    private readonly PdfReadDocument _document;
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<string, List<PdfAccessibilityFinding>> _findings = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _occurrences = new(StringComparer.Ordinal);
    private readonly Dictionary<string, bool> _indeterminate = new(StringComparer.Ordinal);
    private readonly HashSet<string> _applicable = new(StringComparer.Ordinal);
    private readonly HashSet<string> _failed = new(StringComparer.Ordinal);
    private readonly Dictionary<int, int> _pageNumberByObject = new();
    private PdfDictionary? _catalog;
    private bool _catalogHasLanguage;

    private PdfAccessibilityInspector(PdfReadDocument document, CancellationToken cancellationToken) {
        _document = document;
        _objects = document.Objects;
        _cancellationToken = cancellationToken;
    }

    internal static PdfAccessibilityReport Inspect(PdfDocument source, CancellationToken cancellationToken) {
        Guard.NotNull(source, nameof(source));
        var snapshot = source.GetReadSnapshot(cancellationToken: cancellationToken);
        PdfReadDocument document = snapshot.Document;
        var inspector = new PdfAccessibilityInspector(document, cancellationToken);
        int? declaredPart = inspector.Run();
        return inspector.BuildReport(PdfArtifactFingerprint.ComputeSha256(snapshot.Bytes), document.Pages.Count, declaredPart);
    }

    private int? Run() {
        _catalog = _document.CatalogDictionary;
        for (int i = 0; i < _document.Pages.Count; i++) _pageNumberByObject[_document.Pages[i].ObjectNumber] = i + 1;
        string? catalogLanguage = _catalog is null ? null : ReadText(_catalog, "Lang");
        _catalogHasLanguage = !string.IsNullOrEmpty(catalogLanguage);
        if (catalogLanguage is not null) CheckLanguageValue(catalogLanguage, null, null, null, _catalog is null ? null : FindObjectNumber(_catalog));

        if (!PdfPermissionAuthorization.CanExtractText(_document.Security, _document.ReadOptions.PermissionPolicy)) {
            // Neither copying nor accessibility extraction is permitted, so metadata, tags and page content stay unread.
            CheckCatalogFlags();
            CheckEncryption();
            foreach (PdfAccessibilityCondition condition in PdfAccessibilityConditions.All) {
                if (condition.Coverage == PdfAccessibilityCoverage.Automatic && !_applicable.Contains(condition.ConditionId)) {
                    Unknown(condition.ConditionId, null, "The document permissions deny content extraction for accessibility, so this condition was not inspected.");
                }
            }
            AddManualReviewItems();
            return null;
        }

        int? declaredPart = CheckMetadata();
        CheckCatalogFlags();
        CheckOptionalContent();
        CheckEmbeddedFiles();
        CheckXfa();
        CheckEncryption();
        CheckOutlines();
        StructureModel structure = ReadStructure();
        CheckStructure(structure);
        CheckPages(structure);
        AddManualReviewItems();
        return declaredPart;
    }

    private PdfAccessibilityReport BuildReport(string sha256, int pageCount, int? declaredPart) {
        var results = new List<PdfAccessibilityConditionResult>(PdfAccessibilityConditions.All.Count);
        var findings = new List<PdfAccessibilityFinding>();
        foreach (PdfAccessibilityCondition condition in PdfAccessibilityConditions.All) {
            _occurrences.TryGetValue(condition.ConditionId, out int occurrences);
            // Failures are tracked apart from the retained findings so the finding cap never changes a status.
            bool failed = _failed.Contains(condition.ConditionId);
            _findings.TryGetValue(condition.ConditionId, out List<PdfAccessibilityFinding>? list);
            PdfAccessibilityConditionStatus status = condition.Coverage switch {
                PdfAccessibilityCoverage.ManualReview => PdfAccessibilityConditionStatus.NeedsReview,
                PdfAccessibilityCoverage.NotChecked => PdfAccessibilityConditionStatus.NotChecked,
                _ => failed ? PdfAccessibilityConditionStatus.Failed
                    : _indeterminate.ContainsKey(condition.ConditionId) ? PdfAccessibilityConditionStatus.Indeterminate
                    : _applicable.Contains(condition.ConditionId) ? PdfAccessibilityConditionStatus.Passed
                    : PdfAccessibilityConditionStatus.NotApplicable
            };
            results.Add(new PdfAccessibilityConditionResult(condition, status, occurrences));
            if (list is not null) findings.AddRange(list);
        }
        return new PdfAccessibilityReport(sha256, pageCount, declaredPart, results.AsReadOnly(), findings.AsReadOnly());
    }

    private void Applies(string conditionId) => _applicable.Add(conditionId);

    private void Fail(string conditionId, string? rule, string message, int? page = null, int? objectNumber = null, StructNode? element = null) =>
        Add(conditionId, rule, PdfAccessibilityFindingSeverity.Error, message, page, objectNumber, element);

    private void Unknown(string conditionId, string? rule, string message, int? page = null, int? objectNumber = null, StructNode? element = null) {
        _indeterminate[conditionId] = true;
        Add(conditionId, rule, PdfAccessibilityFindingSeverity.Indeterminate, message, page, objectNumber, element);
    }

    private void Add(string conditionId, string? rule, PdfAccessibilityFindingSeverity severity, string message,
        int? page, int? objectNumber, StructNode? element) {
        _applicable.Add(conditionId);
        if (severity == PdfAccessibilityFindingSeverity.Error) _failed.Add(conditionId);
        _occurrences[conditionId] = _occurrences.TryGetValue(conditionId, out int count) ? count + 1 : 1;
        if (!_findings.TryGetValue(conditionId, out List<PdfAccessibilityFinding>? list)) {
            list = new List<PdfAccessibilityFinding>();
            _findings[conditionId] = list;
        }
        if (list.Count >= MaxFindingsPerCondition) return;
        list.Add(new PdfAccessibilityFinding(PdfAccessibilityConditions.Get(conditionId), rule, severity,
            page ?? (element?.PageObjectNumber is int pageObject && _pageNumberByObject.TryGetValue(pageObject, out int number) ? number : null),
            objectNumber, element?.ObjectNumber, element?.Type, message));
    }

    private void AddManualReviewItems() {
        foreach (PdfAccessibilityCondition condition in PdfAccessibilityConditions.All) {
            if (condition.Coverage != PdfAccessibilityCoverage.ManualReview) continue;
            Add(condition.ConditionId, null, PdfAccessibilityFindingSeverity.Review,
                "Review the document for this condition: " + condition.Description, null, null, null);
        }
    }

    // ---- Document level ----------------------------------------------------------------------

    private int? CheckMetadata() {
        Applies("06-001");
        PdfObject? metadataObject = _catalog is not null && _catalog.Items.TryGetValue("Metadata", out PdfObject? value) ? value : null;
        if (Resolve(metadataObject) is not PdfStream) {
            Fail("06-001", "7.1-8", "The catalog has no Metadata stream. Add an XMP packet with dc:title and pdfuaid:part.");
            Fail("06-002", null, "Without XMP metadata the file cannot declare PDF/UA identification.");
            Fail("06-003", null, "Without XMP metadata the file has no dc:title.");
            return null;
        }
        PdfXmpMetadataInfo? xmp;
        try {
            // The accessibility permission was checked by the caller, so read metadata without the copy-permission gate.
            xmp = _document.UncheckedXmpMetadata;
        } catch (Exception exception) when (IsReadFailure(exception)) {
            Unknown("06-002", "5-1", "The XMP metadata stream could not be read: " + exception.Message);
            Unknown("06-003", "7.1-9", "The XMP metadata stream could not be read: " + exception.Message);
            return null;
        }
        Applies("06-002");
        Applies("06-003");
        if (xmp?.PdfUaPart is null) {
            Fail("06-002", "5-1", "The XMP metadata has no pdfuaid:part entry identifying the PDF/UA part.");
        } else if (xmp.PdfUaPart is not (1 or 2)) {
            Fail("06-002", "5-2", "pdfuaid:part is " + xmp.PdfUaPart.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + "; PDF/UA parts are 1 and 2.");
        }
        if (string.IsNullOrWhiteSpace(xmp?.Title)) {
            Fail("06-003", "7.1-9", "The XMP metadata has no dc:title. Add a title that identifies the document.");
        }
        string? raw = xmp?.RawXml;
        if (raw is not null && raw.Contains("x-default")) {
            Applies("11-006");
            if (!_catalogHasLanguage) {
                Fail("11-006", "7.2-33", "The XMP metadata uses x-default language alternatives but the catalog has no Lang entry, so their language cannot be determined.");
            }
        }
        return xmp?.PdfUaPart;
    }

    private void CheckCatalogFlags() {
        Applies("UA1:6.2-1");
        Applies("01-007");
        PdfDictionary? markInfo = _catalog is null ? null : ResolveDict(_catalog.Items.TryGetValue("MarkInfo", out PdfObject? markInfoObject) ? markInfoObject : null);
        if (markInfo is null || Resolve(markInfo.Items.TryGetValue("Marked", out PdfObject? marked) ? marked : null) is not PdfBoolean { Value: true }) {
            Fail("UA1:6.2-1", "6.2-1", "The catalog must contain a MarkInfo dictionary with Marked set to true.");
        }
        if (markInfo is not null && Resolve(markInfo.Items.TryGetValue("Suspects", out PdfObject? suspects) ? suspects : null) is PdfBoolean { Value: true }) {
            Fail("01-007", "7.1-4", "MarkInfo Suspects is true: the producer was unsure the tags are correct. Fix the tags and set Suspects to false.");
        }

        Applies("07-001");
        PdfDictionary? preferences = _catalog is null ? null : ResolveDict(_catalog.Items.TryGetValue("ViewerPreferences", out PdfObject? preferencesObject) ? preferencesObject : null);
        PdfObject? display = preferences is not null && preferences.Items.TryGetValue("DisplayDocTitle", out PdfObject? displayObject) ? Resolve(displayObject) : null;
        if (display is null) {
            Fail("07-001", "7.1-10", "ViewerPreferences has no DisplayDocTitle entry; set it to true so viewers show the title instead of the file name.");
        } else {
            Applies("07-002");
            if (display is not PdfBoolean { Value: true }) {
                Fail("07-002", "7.1-10", "ViewerPreferences DisplayDocTitle is not true.");
            }
        }
    }

    private void CheckOptionalContent() {
        PdfDictionary? properties = _catalog is null ? null : ResolveDict(_catalog.Items.TryGetValue("OCProperties", out PdfObject? value) ? value : null);
        if (properties is null) return;
        if (properties.Items.TryGetValue("D", out PdfObject? defaultObject) && ResolveDict(defaultObject) is PdfDictionary defaultConfig) {
            Applies("20-002");
            Applies("20-003");
            if (string.IsNullOrEmpty(ReadText(defaultConfig, "Name"))) {
                Fail("20-002", "7.10-1", "The default optional content configuration (D) needs a non-empty Name.", objectNumber: ObjectNumberOf(defaultObject));
            }
            if (defaultConfig.Items.ContainsKey("AS")) {
                Fail("20-003", "7.10-2", "The default optional content configuration contains an AS entry.", objectNumber: ObjectNumberOf(defaultObject));
            }
        }
        if (ResolveArray(properties.Items.TryGetValue("Configs", out PdfObject? configsObject) ? configsObject : null) is PdfArray configs) {
            foreach (PdfObject item in configs.Items) {
                if (ResolveDict(item) is not PdfDictionary config) continue;
                Applies("20-001");
                Applies("20-003");
                if (string.IsNullOrEmpty(ReadText(config, "Name"))) {
                    Fail("20-001", "7.10-1", "An optional content configuration in Configs needs a non-empty Name.", objectNumber: ObjectNumberOf(item));
                }
                if (config.Items.ContainsKey("AS")) {
                    Fail("20-003", "7.10-2", "An optional content configuration in Configs contains an AS entry.", objectNumber: ObjectNumberOf(item));
                }
            }
        }
    }

    private void CheckEmbeddedFiles() {
        PdfDictionary? names = _catalog is null ? null : ResolveDict(_catalog.Items.TryGetValue("Names", out PdfObject? namesObject) ? namesObject : null);
        PdfObject? tree = names is not null && names.Items.TryGetValue("EmbeddedFiles", out PdfObject? treeObject) ? treeObject : null;
        if (tree is null) return;
        int nodes = 0;
        var visited = new HashSet<PdfObject>();
        VisitNameTree(tree, 0);

        void VisitNameTree(PdfObject? node, int depth) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (ResolveDict(node) is not PdfDictionary dictionary || !visited.Add(dictionary)) return;
            if (++nodes > _document.ReadOptions.Limits.MaxNameTreeNodes || depth > _document.ReadOptions.Limits.MaxNameTreeDepth) {
                Unknown("21-001", "7.11-1", "The EmbeddedFiles name tree is larger than the read limits allow; remaining entries were not inspected.");
                return;
            }
            if (ResolveArray(dictionary.Items.TryGetValue("Names", out PdfObject? pairs) ? pairs : null) is PdfArray array) {
                for (int i = 1; i < array.Items.Count; i += 2) {
                    CheckFileSpecification(array.Items[i], "21-001", null);
                }
            }
            if (ResolveArray(dictionary.Items.TryGetValue("Kids", out PdfObject? kids) ? kids : null) is PdfArray kidArray) {
                foreach (PdfObject kid in kidArray.Items) VisitNameTree(kid, depth + 1);
            }
        }
    }

    private void CheckFileSpecification(PdfObject? specification, string conditionId, int? page) {
        if (ResolveDict(specification) is not PdfDictionary dictionary || !dictionary.Items.ContainsKey("EF")) return;
        Applies(conditionId);
        bool hasF = !string.IsNullOrEmpty(ReadText(dictionary, "F"));
        bool hasUF = !string.IsNullOrEmpty(ReadText(dictionary, "UF"));
        if (!hasF || !hasUF) {
            Fail(conditionId, "7.11-1", "An embedded file's file specification needs non-empty F and UF entries" +
                (hasF ? " (UF is missing)." : hasUF ? " (F is missing)." : "."), page, ObjectNumberOf(specification));
        }
    }

    private void CheckXfa() {
        PdfDictionary? acroForm = _catalog is null ? null : ResolveDict(_catalog.Items.TryGetValue("AcroForm", out PdfObject? formObject) ? formObject : null);
        if (acroForm is null || !acroForm.Items.TryGetValue("XFA", out PdfObject? xfaObject)) return;
        Applies("25-001");
        var streams = new List<PdfStream>();
        PdfObject? xfa = Resolve(xfaObject);
        if (xfa is PdfStream single) {
            streams.Add(single);
        } else if (xfa is PdfArray packets) {
            for (int i = 1; i < packets.Items.Count; i += 2) {
                string? packet = Resolve(packets.Items[i - 1]) is PdfStringObj name ? name.Value : null;
                if (string.Equals(packet, "config", StringComparison.Ordinal) && Resolve(packets.Items[i]) is PdfStream config) streams.Add(config);
            }
        }
        foreach (PdfStream stream in streams) {
            bool dynamic;
            try {
                byte[] xml = _document.DecodedStreamBudget.DecodeRequired(stream, _objects, _document.ReadOptions.Limits.MaxDecodedStreamBytes, _cancellationToken);
                dynamic = XfaRequiresDynamicRender(xml);
            } catch (Exception exception) when (IsReadFailure(exception) || exception is System.Xml.XmlException) {
                Unknown("25-001", "7.15-1", "The XFA configuration could not be read as XML: " + exception.Message);
                continue;
            }
            if (dynamic) {
                Fail("25-001", "7.15-1", "The XFA configuration sets dynamicRender to required; dynamic XFA forms are not allowed.");
            }
        }
    }

    // Reads the XML itself, so namespace prefixes, encodings and comments cannot change the answer.
    private static bool XfaRequiresDynamicRender(byte[] xml) {
        var settings = new System.Xml.XmlReaderSettings {
            DtdProcessing = System.Xml.DtdProcessing.Prohibit,
            XmlResolver = null,
            ConformanceLevel = System.Xml.ConformanceLevel.Fragment,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true
        };
        using var reader = System.Xml.XmlReader.Create(new MemoryStream(xml, writable: false), settings);
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "dynamicRender", StringComparison.Ordinal) &&
                !reader.IsEmptyElement &&
                string.Equals(reader.ReadElementContentAsString().Trim(), "required", StringComparison.Ordinal)) {
                return true;
            }
        }
        return false;
    }

    private void CheckEncryption() {
        PdfDocumentSecurityInfo security = _document.Security;
        if (!security.HasEncryption) return;
        Applies("26-001");
        if (security.EncryptionPermissions is not int permissions) {
            Fail("26-001", "7.16-1", "The encryption dictionary has no P entry.", objectNumber: security.EncryptObjectNumber);
            return;
        }
        Applies("26-002");
        if ((permissions & (1 << 9)) == 0) {
            Fail("26-002", "7.16-1", "The P entry turns off bit 10, which blocks text extraction for accessibility.", objectNumber: security.EncryptObjectNumber);
        }
    }

    private void CheckOutlines() {
        PdfDictionary? outlines = _catalog is null ? null : ResolveDict(_catalog.Items.TryGetValue("Outlines", out PdfObject? value) ? value : null);
        if (outlines is null || !outlines.Items.ContainsKey("First")) return;
        Applies("11-003");
        if (!_catalogHasLanguage) {
            Fail("11-003", "7.2-2", "The document has bookmarks but the catalog has no Lang entry, so their language cannot be determined.");
        }
    }

    private void CheckLanguageValue(string value, int? page, StructNode? element, string? where, int? objectNumber) {
        Applies("UA1:7.2-29");
        if (!LanguagePattern.IsMatch(value)) {
            Fail("UA1:7.2-29", "7.2-29", "'" + value + "' is not a valid language identifier" + (where is null ? "." : " in " + where + "."), page, objectNumber, element);
        }
    }

    // ---- Object helpers ----------------------------------------------------------------------

    private static bool IsReadFailure(Exception exception) =>
        exception is PdfReadLimitException or PdfParseException or InvalidDataException or FormatException or NotSupportedException or IOException;

    private PdfObject? Resolve(PdfObject? value) => PdfObjectLookup.ResolveChainCancellable(_objects, value, _cancellationToken);

    private PdfDictionary? ResolveDict(PdfObject? value) => Resolve(value) switch {
        PdfDictionary dictionary => dictionary,
        PdfStream stream => stream.Dictionary,
        _ => null
    };

    private PdfArray? ResolveArray(PdfObject? value) => Resolve(value) as PdfArray;

    private static int? ObjectNumberOf(PdfObject? value) => value is PdfReference reference ? reference.ObjectNumber : null;

    private string? ReadText(PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) && Resolve(value) is PdfStringObj text ? text.Value : null;

    private string? ReadName(PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) && Resolve(value) is PdfName name ? name.Name : null;

    private int? ReadInteger(PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) && Resolve(value) is PdfNumber number &&
        number.Value >= int.MinValue && number.Value <= int.MaxValue ? (int)number.Value : null;

    private int? FindObjectNumber(PdfDictionary dictionary) {
        foreach (KeyValuePair<int, PdfIndirectObject> item in _objects) {
            if (item.Value.IsValueBuilt && ReferenceEquals(item.Value.Value, dictionary)) return item.Key;
        }
        return null;
    }
}
