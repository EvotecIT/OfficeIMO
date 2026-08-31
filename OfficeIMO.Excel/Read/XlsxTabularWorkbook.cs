#nullable enable

using OfficeIMO.Core.Internal;
using System.Data.Common;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Owns the minimal validated package state needed by the forward-only XLSX reader.
    /// Unsupported package shapes route back to <see cref="ExcelDocumentReader"/>.
    /// </summary>
    internal sealed class XlsxTabularWorkbook : IDisposable {
        private const string PackageRelationshipsNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string PackageContentTypesNamespace =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string TransitionalOfficeRelationshipsNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string StrictOfficeRelationshipsNamespace =
            "http://purl.oclc.org/ooxml/officeDocument/relationships";
        private const string TransitionalSpreadsheetNamespace =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string StrictSpreadsheetNamespace =
            "http://purl.oclc.org/ooxml/spreadsheetml/main";
        private const string WorksheetRelationshipSuffix = "/worksheet";
        private const string ChartSheetRelationshipSuffix = "/chartsheet";
        private const string DialogSheetRelationshipSuffix = "/dialogsheet";
        private const string MacroSheetRelationshipSuffix = "/macrosheet";
        private const string InternationalMacroSheetRelationshipSuffix = "/intlMacrosheet";
        private const string SharedStringsRelationshipSuffix = "/sharedStrings";
        private const string StylesRelationshipSuffix = "/styles";
        private const int MaximumPrefetchedWorksheetBytes = 64 * 1024 * 1024;
        private const string WorksheetContentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string SharedStringsContentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        private const string StylesContentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

        private static readonly HashSet<string> SupportedWorkbookContentTypes = new(StringComparer.OrdinalIgnoreCase) {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
            "application/vnd.ms-excel.addin.macroEnabled.main+xml"
        };

        private readonly OpenXmlPackagePartBufferReader _parts;
        private readonly OpenXmlPackagePartBufferReader? _prefetchedParts;
        private readonly string? _prefetchedSheetPartName;
        private readonly IDisposable? _ownedResource;
        private readonly SharedStringCache _sharedStrings;
        private readonly StylesCacheProvider _styles;
        private readonly ExcelReadOptions _options;
        private readonly Dictionary<string, string> _contentTypeOverrides;
        private readonly Dictionary<string, string> _contentTypeDefaults;
        private readonly XlsxTabularSheet[] _sheets;
        private readonly string[] _tableNames;
        private bool _disposed;

        private XlsxTabularWorkbook(
            OpenXmlPackagePartBufferReader parts,
            OpenXmlPackagePartBufferReader? prefetchedParts,
            IDisposable? ownedResource,
            ExcelReadOptions options,
            bool metadataOnly = false) {
            _parts = parts;
            _prefetchedParts = prefetchedParts;
            _ownedResource = ownedResource;
            _options = options;
            options.CancellationToken.ThrowIfCancellationRequested();

            (_contentTypeOverrides, _contentTypeDefaults) = ReadContentTypes();
            string workbookPartName = ReadWorkbookPartName();
            ValidatePartContentType(
                workbookPartName,
                SupportedWorkbookContentTypes,
                "workbook");
            IReadOnlyDictionary<string, PackageRelationship> workbookRelationships =
                ReadRelationships(workbookPartName);
            (_sheets, ExcelDateSystem dateSystem) = ReadWorkbook(
                workbookPartName,
                workbookRelationships,
                metadataOnly);
            if (_sheets.Length == 0) {
                if (metadataOnly) {
                    throw new InvalidDataException("The workbook contains no readable worksheets.");
                }
                throw new XlsxTabularFastPathNotSupportedException(
                    "The workbook contains no native-path worksheets.");
            }

            DateSystem = dateSystem;
            _tableNames = _sheets.Select(static sheet => sheet.Name).ToArray();
            if (metadataOnly) {
                _prefetchedSheetPartName = null;
                _sharedStrings = SharedStringCache.Empty(options);
                _styles = new StylesCacheProvider(StylesCache.Empty());
                options.CancellationToken.ThrowIfCancellationRequested();
                return;
            }

            XlsxTabularSheet? prefetchedSheet = options.EnableWorksheetPrefetch
                ? ResolvePrefetchedSheet(_sheets, options)
                : null;
            _prefetchedSheetPartName = prefetchedSheet?.PartName;
            if (_prefetchedSheetPartName != null && _prefetchedParts != null) {
                _prefetchedParts.BeginPrefetch(
                    _prefetchedSheetPartName,
                    MaximumPrefetchedWorksheetBytes,
                    options.CancellationToken);
            }
            int maximumPartBytes = options.MaxInputBytes > int.MaxValue
                ? int.MaxValue
                : checked((int)options.MaxInputBytes);

            string? sharedStringsPart = ResolveOptionalPart(
                workbookPartName,
                workbookRelationships,
                SharedStringsRelationshipSuffix,
                "shared-string table",
                SharedStringsContentType);
            _sharedStrings = sharedStringsPart == null
                ? SharedStringCache.Empty(options)
                : SharedStringCache.Build(
                    () => _parts.OpenPart(sharedStringsPart, maximumPartBytes),
                    options);

            string? stylesPart = ResolveOptionalPart(
                workbookPartName,
                workbookRelationships,
                StylesRelationshipSuffix,
                "styles",
                StylesContentType);
            _styles = stylesPart == null
                ? new StylesCacheProvider(StylesCache.Empty())
                : new StylesCacheProvider(
                    () => _parts.OpenPart(stylesPart, maximumPartBytes));
        }

        internal IReadOnlyList<string> TableNames => _tableNames;

        internal ExcelDateSystem DateSystem { get; }

        internal static XlsxTabularWorkbook Open(string path, ExcelReadOptions options) {
            return Open(path, options, metadataOnly: false);
        }

        internal static IReadOnlyList<string> ReadSheetNames(string path, ExcelReadOptions options) {
            using XlsxTabularWorkbook workbook = Open(path, options, metadataOnly: true);
            options.CancellationToken.ThrowIfCancellationRequested();
            return workbook._tableNames.ToArray();
        }

        private static XlsxTabularWorkbook Open(
            string path,
            ExcelReadOptions options,
            bool metadataOnly) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("File path cannot be empty.", nameof(path));
            }
            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            SharedReadOnlyFileSnapshot? snapshot = null;
            OpenXmlPackagePartBufferReader? parts = null;
            OpenXmlPackagePartBufferReader? prefetchedParts = null;
            try {
                snapshot = SharedReadOnlyFileSnapshot.Open(path);
                if (snapshot.Length > options.MaxInputBytes) {
                    throw new InvalidDataException(
                        $"Workbook input contains {snapshot.Length} bytes, exceeding the configured limit of {options.MaxInputBytes} bytes.");
                }

                parts = OpenXmlPackagePartBufferReader.TryOpen(snapshot.CreateView(bufferSize: 1))
                    ?? throw new XlsxTabularFastPathNotSupportedException(
                        "The workbook is not a readable Open XML package.");
                if (!metadataOnly && options.EnableWorksheetPrefetch) {
                    prefetchedParts = OpenXmlPackagePartBufferReader.TryOpen(snapshot.CreateView(bufferSize: 1));
                }
                var workbook = new XlsxTabularWorkbook(parts, prefetchedParts, snapshot, options, metadataOnly);
                parts = null;
                prefetchedParts = null;
                snapshot = null;
                return workbook;
            } catch {
                parts?.Dispose();
                prefetchedParts?.Dispose();
                snapshot?.Dispose();
                throw;
            }
        }

        internal static XlsxTabularWorkbook Open(byte[] bytes, ExcelReadOptions options) {
            if (bytes == null) {
                throw new ArgumentNullException(nameof(bytes));
            }
            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }
            if (bytes.LongLength > options.MaxInputBytes) {
                throw new InvalidDataException(
                    $"Workbook input contains {bytes.LongLength} bytes, exceeding the configured limit of {options.MaxInputBytes} bytes.");
            }

            OpenXmlPackagePartBufferReader parts = OpenXmlPackagePartBufferReader.TryOpen(bytes)
                ?? throw new XlsxTabularFastPathNotSupportedException(
                    "The workbook is not a readable Open XML package.");
            OpenXmlPackagePartBufferReader? prefetchedParts = null;
            try {
                prefetchedParts = options.EnableWorksheetPrefetch
                    ? OpenXmlPackagePartBufferReader.TryOpen(bytes)
                    : null;
                var workbook = new XlsxTabularWorkbook(parts, prefetchedParts, ownedResource: null, options);
                parts = null!;
                prefetchedParts = null;
                return workbook;
            } catch {
                parts.Dispose();
                prefetchedParts?.Dispose();
                throw;
            }
        }

        internal DbDataReader OpenTable(
            string tableName,
            bool hasHeaderRow,
            CancellationToken cancellationToken) {
            ThrowIfDisposed();
            XlsxTabularSheet? sheet = _sheets.FirstOrDefault(
                candidate => string.Equals(candidate.Name, tableName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null) {
                throw new KeyNotFoundException($"Worksheet '{tableName}' was not found.");
            }

            var reader = new ExcelSheetReader(
                sheet.Name,
                sheet.PartName,
                _sharedStrings,
                _styles,
                _options,
                DateSystem,
                string.Equals(sheet.PartName, _prefetchedSheetPartName, StringComparison.OrdinalIgnoreCase)
                    ? _prefetchedParts ?? _parts
                    : _parts);
            DbDataReader dataReader = string.IsNullOrWhiteSpace(_options.A1Range)
                ? (DbDataReader)reader.ReadUsedRangeAsDataReader(
                    hasHeaderRow,
                    schemaSampleRows: 0,
                    cancellationToken)
                : (DbDataReader)reader.ReadRangeAsDataReader(
                    _options.A1Range!,
                    hasHeaderRow,
                    chunkRows: Math.Min(1024, _options.MaxDataReaderChunkRows),
                    schemaSampleRows: 0,
                    ct: cancellationToken);

            return _options.InferSchema && _options.SchemaSampleRows > 0
                ? ExcelSchemaInferenceDataReader.Create(
                    dataReader,
                    _options.SchemaSampleRows,
                    _options.MaxDataReaderSchemaSampleRows,
                    _options.MaxDataReaderBufferedCells,
                    _options.Culture,
                    cancellationToken)
                : dataReader;
        }

        private string ReadWorkbookPartName() {
            IReadOnlyDictionary<string, PackageRelationship> relationships =
                ReadRelationships(string.Empty);
            PackageRelationship[] candidates = relationships.Values
                .Where(relationship => IsOfficeRelationship(
                    relationship.Type,
                    "/officeDocument"))
                .ToArray();
            if (candidates.Length != 1 || candidates[0].IsExternal) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package does not contain one internal Office workbook relationship.");
            }

            string workbookPartName = ResolveTarget(string.Empty, candidates[0].Target);
            if (!_parts.ContainsPart(workbookPartName)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package workbook relationship target is missing.");
            }

            return workbookPartName;
        }

        private (Dictionary<string, string> Overrides, Dictionary<string, string> Defaults)
            ReadContentTypes() {
            XDocument contentTypes = ReadXmlPart("[Content_Types].xml", _options.MaxMetadataPartBytes);
            XNamespace ns = PackageContentTypesNamespace;
            if (contentTypes.Root?.Name != ns + "Types") {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package content-type manifest namespace is not supported by the native path.");
            }

            var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement element in contentTypes.Root.Elements()) {
                _options.CancellationToken.ThrowIfCancellationRequested();
                if (element.Name == ns + "Override") {
                    string? rawPartName = (string?)element.Attribute("PartName");
                    string? contentType = (string?)element.Attribute("ContentType");
                    string normalizedPartName = NormalizeContentTypePartName(rawPartName);
                    if (normalizedPartName.Length == 0
                        || rawPartName![0] != '/'
                        || !IsValidContentType(contentType)
                        || overrides.ContainsKey(normalizedPartName)) {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "The package content-type overrides require the Open XML SDK fallback path.");
                    }
                    overrides.Add(normalizedPartName, contentType!);
                    continue;
                }

                if (element.Name == ns + "Default") {
                    string? extension = (string?)element.Attribute("Extension");
                    string? contentType = (string?)element.Attribute("ContentType");
                    if (!IsValidContentTypeExtension(extension)
                        || !IsValidContentType(contentType)
                        || defaults.ContainsKey(extension!)) {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "The package content-type defaults require the Open XML SDK fallback path.");
                    }
                    defaults.Add(extension!, contentType!);
                    continue;
                }

                throw new XlsxTabularFastPathNotSupportedException(
                    "The package content-type manifest requires the Open XML SDK fallback path.");
            }

            return (overrides, defaults);
        }

        private void ValidatePartContentType(
            string partName,
            ISet<string> supportedContentTypes,
            string role) {
            string? contentType = GetPartContentType(partName);
            if (string.IsNullOrWhiteSpace(contentType)
                || !supportedContentTypes.Contains(contentType!)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    $"The {role} content type is not supported by the native path.");
            }
        }

        private void ValidatePartContentType(string partName, string expectedContentType, string role) {
            string? contentType = GetPartContentType(partName);
            if (!string.Equals(contentType, expectedContentType, StringComparison.OrdinalIgnoreCase)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    $"The {role} content type is not supported by the native path.");
            }
        }

        private string? GetPartContentType(string partName) {
            string expectedPartName = "/" + partName.TrimStart('/');
            if (_contentTypeOverrides.TryGetValue(expectedPartName, out string? contentType)) {
                return contentType;
            }

            int extensionSeparator = partName.LastIndexOf('.');
            if (extensionSeparator < 0 || extensionSeparator == partName.Length - 1) {
                return null;
            }

            string extension = partName.Substring(extensionSeparator + 1);
            return _contentTypeDefaults.TryGetValue(extension, out contentType)
                ? contentType
                : null;
        }

        private IReadOnlyDictionary<string, PackageRelationship> ReadRelationships(
            string sourcePartName) {
            string relationshipPartName = GetRelationshipPartName(sourcePartName);
            XDocument relationships = ReadXmlPart(
                relationshipPartName,
                _options.MaxMetadataPartBytes);
            XNamespace ns = PackageRelationshipsNamespace;
            if (relationships.Root?.Name != ns + "Relationships") {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package relationship namespace is not supported by the native path.");
            }

            var result = new Dictionary<string, PackageRelationship>(StringComparer.Ordinal);
            foreach (XElement element in relationships.Root.Elements()) {
                _options.CancellationToken.ThrowIfCancellationRequested();
                if (element.Name != ns + "Relationship") {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The package relationships require the Open XML SDK fallback path.");
                }

                string? id = (string?)element.Attribute("Id");
                string? type = (string?)element.Attribute("Type");
                string? target = (string?)element.Attribute("Target");
                if (string.IsNullOrWhiteSpace(id)
                    || string.IsNullOrWhiteSpace(type)
                    || string.IsNullOrWhiteSpace(target)
                    || !IsValidRelationshipId(id!)
                    || !Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out _)
                    || result.ContainsKey(id!)) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The package relationships require the Open XML SDK fallback path.");
                }

                result.Add(
                    id!,
                    new PackageRelationship(
                        type!,
                        target!,
                        ReadRelationshipTargetMode(element)));
            }

            return result;
        }

        private (XlsxTabularSheet[] Sheets, ExcelDateSystem DateSystem) ReadWorkbook(
            string workbookPartName,
            IReadOnlyDictionary<string, PackageRelationship> relationships,
            bool metadataOnly) {
            XDocument workbook = ReadXmlPart(workbookPartName, _options.MaxMetadataPartBytes);
            if (workbook.Root == null
                || (workbook.Root.Name.NamespaceName != TransitionalSpreadsheetNamespace
                    && workbook.Root.Name.NamespaceName != StrictSpreadsheetNamespace)
                || workbook.Root.Name.LocalName != "workbook") {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The workbook XML namespace is not supported by the native path.");
            }

            XNamespace spreadsheet = workbook.Root.Name.Namespace;
            ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred;
            XElement? workbookProperties = workbook.Root.Element(spreadsheet + "workbookPr");
            string? date1904 = (string?)workbookProperties?.Attribute("date1904");
            if (!string.IsNullOrEmpty(date1904)) {
                if (date1904 == "1" || string.Equals(date1904, "true", StringComparison.OrdinalIgnoreCase)) {
                    dateSystem = ExcelDateSystem.NineteenFour;
                } else if (date1904 != "0" && !string.Equals(date1904, "false", StringComparison.OrdinalIgnoreCase)) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The workbook date-system flag requires the Open XML SDK fallback path.");
                }
            }

            XElement? sheetsElement = workbook.Root.Element(spreadsheet + "sheets");
            if (sheetsElement == null) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The workbook has no sheets collection.");
            }

            var workbookSheets = new List<XElement>();
            foreach (XElement sheet in sheetsElement.Elements(spreadsheet + "sheet")) {
                _options.CancellationToken.ThrowIfCancellationRequested();
                if (workbookSheets.Count >= _options.MaxWorksheets) {
                    throw new InvalidDataException(
                        $"The workbook contains more than the configured {_options.MaxWorksheets} worksheet definitions.");
                }
                workbookSheets.Add(sheet);
            }
            if (string.IsNullOrWhiteSpace(_options.SheetName)
                && !_options.SheetIndex.HasValue
                && workbookSheets.Count > 1
                && !metadataOnly) {
                // The public multi-result reader still uses the SDK path. Stop before
                // resolving every sheet and optional global part so that fallback does
                // not pay the complete native metadata probe first.
                throw new XlsxTabularFastPathNotSupportedException(
                    "Multi-result XLSX reads retain the complete Open XML SDK path.");
            }

            XNamespace transitionalRelationships = TransitionalOfficeRelationshipsNamespace;
            XNamespace strictRelationships = StrictOfficeRelationshipsNamespace;
            var sheets = new List<XlsxTabularSheet>();
            var worksheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement sheet in workbookSheets) {
                _options.CancellationToken.ThrowIfCancellationRequested();
                string? name = (string?)sheet.Attribute("name");
                string? relationshipId = (string?)sheet.Attribute(transitionalRelationships + "id")
                    ?? (string?)sheet.Attribute(strictRelationships + "id");
                if (string.IsNullOrEmpty(name)
                    || string.IsNullOrEmpty(relationshipId)
                    || !relationships.TryGetValue(relationshipId!, out PackageRelationship? relationship)) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "A workbook sheet requires the Open XML SDK fallback path.");
                }
                if (relationship.IsExternal) {
                    throw new InvalidDataException(
                        $"The OpenXML worksheet '{name}' references external relationship '{relationshipId}'.");
                }
                if (!IsOfficeRelationship(relationship.Type, WorksheetRelationshipSuffix)) {
                    if (IsSupportedNonWorksheetRelationship(relationship.Type)) {
                        if (metadataOnly) {
                            continue;
                        }
                        throw new XlsxTabularFastPathNotSupportedException(
                            "Non-worksheet sheet relationships require the Open XML SDK fallback path.");
                    }

                    throw new XlsxTabularFastPathNotSupportedException(
                        "A workbook sheet relationship requires the Open XML SDK fallback path.");
                }
                if (!worksheetNames.Add(name!)) {
                    throw new InvalidDataException(
                        $"The workbook contains duplicate worksheet name '{name}' under case-insensitive matching.");
                }

                string partName = ResolveTarget(workbookPartName, relationship.Target);
                if (!_parts.ContainsPart(partName)) {
                    throw new InvalidDataException(
                        $"The OpenXML worksheet '{name}' references missing relationship '{relationshipId}'.");
                }
                ValidatePartContentType(partName, WorksheetContentType, "worksheet");
                sheets.Add(new XlsxTabularSheet(name!, partName));
            }

            _options.CancellationToken.ThrowIfCancellationRequested();
            return (sheets.ToArray(), dateSystem);
        }

        private string? ResolveOptionalPart(
            string workbookPartName,
            IReadOnlyDictionary<string, PackageRelationship> relationships,
            string relationshipSuffix,
            string relationshipName,
            string expectedContentType) {
            PackageRelationship[] matches = relationships.Values
                .Where(relationship => IsOfficeRelationship(
                    relationship.Type,
                    relationshipSuffix))
                .Take(2)
                .ToArray();
            if (matches.Length == 0) {
                return null;
            }
            if (matches.Length != 1 || matches[0].IsExternal) {
                throw new XlsxTabularFastPathNotSupportedException(
                    $"The workbook {relationshipName} relationship requires the Open XML SDK fallback path.");
            }

            string partName = ResolveTarget(workbookPartName, matches[0].Target);
            if (!_parts.ContainsPart(partName)) {
                throw new InvalidDataException(
                    $"The workbook {relationshipName} part '{partName}' is missing.");
            }
            ValidatePartContentType(partName, expectedContentType, relationshipName);

            return partName;
        }

        private XDocument ReadXmlPart(string partName, int maximumBytes) {
            try {
                using Stream stream = _parts.OpenPart(partName, maximumBytes);
                using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    CloseInput = false,
                    MaxCharactersInDocument = maximumBytes
                });
                XDocument document = XDocument.Load(reader, LoadOptions.None);
                _options.CancellationToken.ThrowIfCancellationRequested();
                return document;
            } catch (XmlException exception) {
                throw new XlsxTabularFastPathNotSupportedException(
                    $"Package part '{partName}' requires the Open XML SDK fallback path.",
                    exception);
            }
        }

        private static string ResolveTarget(string sourcePartName, string target) {
            if (string.IsNullOrWhiteSpace(target)
                || target.IndexOf('\\') >= 0
                || target.IndexOf('?') >= 0
                || target.IndexOf('#') >= 0
                || Uri.TryCreate(target, UriKind.Absolute, out _)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "A package relationship target is not supported by the native path.");
            }

            string source = sourcePartName.TrimStart('/');
            int separator = source.LastIndexOf('/');
            string directory = separator < 0 ? string.Empty : source.Substring(0, separator + 1);
            string combined = target.StartsWith("/", StringComparison.Ordinal)
                ? target.TrimStart('/')
                : directory + target;
            var segments = new List<string>();
            foreach (string encodedSegment in combined.Split('/')) {
                if (encodedSegment.Length == 0) {
                    continue;
                }
                string segment = DecodePartSegment(encodedSegment);
                if (segment == ".") continue;
                if (segment == "..") {
                    if (segments.Count == 0) {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "A package relationship target escapes the package root.");
                    }
                    segments.RemoveAt(segments.Count - 1);
                    continue;
                }
                segments.Add(segment);
            }

            if (segments.Count == 0) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "A package relationship target does not identify a part.");
            }

            return string.Join("/", segments);
        }

        private static string GetRelationshipPartName(string sourcePartName) {
            string source = sourcePartName.TrimStart('/');
            int separator = source.LastIndexOf('/');
            string directory = separator < 0 ? string.Empty : source.Substring(0, separator + 1);
            string fileName = separator < 0 ? source : source.Substring(separator + 1);
            return directory + "_rels/" + fileName + ".rels";
        }

        private static string NormalizeContentTypePartName(string? partName) {
            if (string.IsNullOrWhiteSpace(partName)
                || !partName!.StartsWith("/", StringComparison.Ordinal)
                || partName.StartsWith("//", StringComparison.Ordinal)
                || partName.EndsWith("/", StringComparison.Ordinal)
                || partName!.IndexOf('\\') >= 0
                || partName.IndexOf('?') >= 0
                || partName.IndexOf('#') >= 0) {
                return string.Empty;
            }

            string[] encodedSegments = partName.Split('/');
            if (encodedSegments.Length < 2) {
                return string.Empty;
            }

            var segments = new List<string>(encodedSegments.Length - 1);
            try {
                foreach (string encodedSegment in encodedSegments.Skip(1)) {
                    if (encodedSegment.Length == 0) return string.Empty;
                    string segment = DecodePartSegment(encodedSegment);
                    if (segment == "." || segment == "..") return string.Empty;
                    segments.Add(segment);
                }
            } catch (XlsxTabularFastPathNotSupportedException) {
                return string.Empty;
            }

            return "/" + string.Join("/", segments);
        }

        private static string DecodePartSegment(string segment) {
            for (int index = 0; index < segment.Length; index++) {
                if (segment[index] != '%') continue;
                if (index + 2 >= segment.Length
                    || !IsHexDigit(segment[index + 1])
                    || !IsHexDigit(segment[index + 2])) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "A package part URI contains invalid percent encoding.");
                }
                index += 2;
            }

            string decoded;
            try {
                decoded = Uri.UnescapeDataString(segment);
            } catch (UriFormatException exception) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "A package part URI contains invalid percent encoding.",
                    exception);
            }
            if (decoded.Length == 0
                || decoded.IndexOf('/') >= 0
                || decoded.IndexOf('\\') >= 0
                || decoded.Any(char.IsControl)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "A package part URI contains an unsafe encoded segment.");
            }

            return decoded;
        }

        private static bool IsHexDigit(char value) =>
            value is >= '0' and <= '9'
            || value is >= 'a' and <= 'f'
            || value is >= 'A' and <= 'F';

        private static bool IsValidContentType(string? contentType) {
            if (string.IsNullOrWhiteSpace(contentType)
                || !string.Equals(contentType, contentType!.Trim(), StringComparison.Ordinal)) {
                return false;
            }

            int separator = contentType.IndexOf('/');
            return separator > 0
                && separator == contentType.LastIndexOf('/')
                && separator < contentType.Length - 1
                && IsContentTypeToken(contentType, 0, separator)
                && IsContentTypeToken(contentType, separator + 1, contentType.Length);
        }

        private static bool IsContentTypeToken(string contentType, int start, int end) {
            for (int index = start; index < end; index++) {
                if (!IsContentTypeTokenCharacter(contentType[index])) {
                    return false;
                }
            }
            return true;
        }

        private static bool IsContentTypeTokenCharacter(char character) =>
            character is >= 'a' and <= 'z'
            || character is >= 'A' and <= 'Z'
            || character is >= '0' and <= '9'
            || character is '!' or '#' or '$' or '%' or '&' or '\'' or '*'
                or '+' or '-' or '.' or '^' or '_' or '`' or '|' or '~';

        private static bool IsValidContentTypeExtension(string? extension) =>
            !string.IsNullOrWhiteSpace(extension)
            && extension!.All(static character =>
                character is >= 'a' and <= 'z'
                || character is >= 'A' and <= 'Z'
                || character is >= '0' and <= '9'
                || character is '-' or '_');

        private static bool IsValidRelationshipId(string id) {
            try {
                XmlConvert.VerifyNCName(id);
                return true;
            } catch (XmlException) {
                return false;
            }
        }

        private static bool ReadRelationshipTargetMode(XElement relationship) {
            XAttribute? attribute = relationship.Attribute("TargetMode");
            if (attribute == null || attribute.Value == "Internal") {
                return false;
            }
            if (attribute.Value == "External") {
                return true;
            }

            throw new XlsxTabularFastPathNotSupportedException(
                "A package relationship target mode requires the Open XML SDK fallback path.");
        }

        private static bool IsSupportedNonWorksheetRelationship(string relationshipType) =>
            IsOfficeRelationship(relationshipType, ChartSheetRelationshipSuffix)
            || IsOfficeRelationship(relationshipType, DialogSheetRelationshipSuffix)
            || IsOfficeRelationship(relationshipType, MacroSheetRelationshipSuffix)
            || IsOfficeRelationship(relationshipType, InternationalMacroSheetRelationshipSuffix);

        private static bool IsOfficeRelationship(string? relationshipType, string suffix) =>
            string.Equals(
                relationshipType,
                TransitionalOfficeRelationshipsNamespace + suffix,
                StringComparison.Ordinal)
            || string.Equals(
                relationshipType,
                StrictOfficeRelationshipsNamespace + suffix,
                StringComparison.Ordinal);

        private static XlsxTabularSheet? ResolvePrefetchedSheet(
            IReadOnlyList<XlsxTabularSheet> sheets,
            ExcelReadOptions options) {
            if (!string.IsNullOrWhiteSpace(options.SheetName)) {
                return sheets.FirstOrDefault(sheet => string.Equals(
                    sheet.Name,
                    options.SheetName,
                    StringComparison.OrdinalIgnoreCase));
            }
            if (options.SheetIndex is int sheetIndex) {
                return (uint)sheetIndex < (uint)sheets.Count ? sheets[sheetIndex] : null;
            }
            return sheets.Count == 0 ? null : sheets[0];
        }

        private void ThrowIfDisposed() {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(XlsxTabularWorkbook));
            }
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }

            _disposed = true;
            try {
                try {
                    _prefetchedParts?.Dispose();
                } finally {
                    _parts.Dispose();
                }
            } finally {
                _ownedResource?.Dispose();
            }
        }

        private sealed class PackageRelationship {
            internal PackageRelationship(string type, string target, bool isExternal) {
                Type = type;
                Target = target;
                IsExternal = isExternal;
            }

            internal string Type { get; }

            internal string Target { get; }

            internal bool IsExternal { get; }
        }

        private sealed class XlsxTabularSheet {
            internal XlsxTabularSheet(string name, string partName) {
                Name = name;
                PartName = partName;
            }

            internal string Name { get; }

            internal string PartName { get; }
        }
    }
}
