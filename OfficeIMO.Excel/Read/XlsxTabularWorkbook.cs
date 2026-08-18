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
        private const int MaximumMetadataPartBytes = 16 * 1024 * 1024;

        private static readonly HashSet<string> SupportedWorkbookContentTypes = new(StringComparer.OrdinalIgnoreCase) {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
            "application/vnd.ms-excel.addin.macroEnabled.main+xml"
        };

        private readonly OpenXmlPackagePartBufferReader _parts;
        private readonly IDisposable? _ownedResource;
        private readonly SharedStringCache _sharedStrings;
        private readonly StylesCacheProvider _styles;
        private readonly ExcelReadOptions _options;
        private readonly XlsxTabularSheet[] _sheets;
        private readonly string[] _tableNames;
        private bool _disposed;

        private XlsxTabularWorkbook(
            OpenXmlPackagePartBufferReader parts,
            IDisposable? ownedResource,
            ExcelReadOptions options) {
            _parts = parts;
            _ownedResource = ownedResource;
            _options = options;
            options.CancellationToken.ThrowIfCancellationRequested();

            string workbookPartName = ReadWorkbookPartName();
            ValidateWorkbookContentType(workbookPartName);
            IReadOnlyDictionary<string, PackageRelationship> workbookRelationships =
                ReadRelationships(workbookPartName);
            (_sheets, ExcelDateSystem dateSystem) = ReadWorkbook(
                workbookPartName,
                workbookRelationships);
            if (_sheets.Length == 0) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The workbook contains no native-path worksheets.");
            }

            DateSystem = dateSystem;
            _tableNames = _sheets.Select(static sheet => sheet.Name).ToArray();
            int maximumPartBytes = options.MaxInputBytes > int.MaxValue
                ? int.MaxValue
                : checked((int)options.MaxInputBytes);

            string? sharedStringsPart = ResolveOptionalPart(
                workbookPartName,
                workbookRelationships,
                SharedStringsRelationshipSuffix,
                "shared-string table");
            _sharedStrings = sharedStringsPart == null
                ? SharedStringCache.Empty(options)
                : SharedStringCache.Build(
                    () => _parts.OpenPart(sharedStringsPart, maximumPartBytes),
                    options);

            string? stylesPart = ResolveOptionalPart(
                workbookPartName,
                workbookRelationships,
                StylesRelationshipSuffix,
                "styles");
            _styles = stylesPart == null
                ? new StylesCacheProvider(StylesCache.Empty())
                : new StylesCacheProvider(
                    () => _parts.OpenPart(stylesPart, maximumPartBytes));
        }

        internal IReadOnlyList<string> TableNames => _tableNames;

        internal ExcelDateSystem DateSystem { get; }

        internal static XlsxTabularWorkbook Open(string path, ExcelReadOptions options) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("File path cannot be empty.", nameof(path));
            }
            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            SharedReadOnlyFileSnapshot? snapshot = null;
            OpenXmlPackagePartBufferReader? parts = null;
            try {
                snapshot = SharedReadOnlyFileSnapshot.Open(path);
                if (snapshot.Length > options.MaxInputBytes) {
                    throw new InvalidDataException(
                        $"Workbook input contains {snapshot.Length} bytes, exceeding the configured limit of {options.MaxInputBytes} bytes.");
                }

                parts = OpenXmlPackagePartBufferReader.TryOpen(snapshot.CreateView(bufferSize: 1))
                    ?? throw new XlsxTabularFastPathNotSupportedException(
                        "The workbook is not a readable Open XML package.");
                var workbook = new XlsxTabularWorkbook(parts, snapshot, options);
                parts = null;
                snapshot = null;
                return workbook;
            } catch {
                parts?.Dispose();
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
            try {
                var workbook = new XlsxTabularWorkbook(parts, ownedResource: null, options);
                parts = null!;
                return workbook;
            } catch {
                parts.Dispose();
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
                _parts);
            return (DbDataReader)reader.ReadUsedRangeAsDataReader(
                hasHeaderRow,
                schemaSampleRows: 0,
                cancellationToken);
        }

        private string ReadWorkbookPartName() {
            XDocument relationships = ReadXmlPart("_rels/.rels", MaximumMetadataPartBytes);
            XNamespace ns = PackageRelationshipsNamespace;
            if (relationships.Root?.Name != ns + "Relationships") {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package root relationship namespace is not supported by the native path.");
            }

            XElement[] candidates = relationships.Root
                .Elements(ns + "Relationship")
                .Where(element =>
                    IsOfficeRelationship(
                        (string?)element.Attribute("Type"),
                        "/officeDocument"))
                .ToArray();
            if (candidates.Length != 1
                || string.Equals(
                    (string?)candidates[0].Attribute("TargetMode"),
                    "External",
                    StringComparison.OrdinalIgnoreCase)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package does not contain one internal Office workbook relationship.");
            }

            string? target = (string?)candidates[0].Attribute("Target");
            if (string.IsNullOrWhiteSpace(target)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package workbook relationship has no target.");
            }

            string workbookPartName = ResolveTarget(string.Empty, target!);
            if (!_parts.ContainsPart(workbookPartName)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package workbook relationship target is missing.");
            }

            return workbookPartName;
        }

        private void ValidateWorkbookContentType(string workbookPartName) {
            XDocument contentTypes = ReadXmlPart("[Content_Types].xml", MaximumMetadataPartBytes);
            XNamespace ns = PackageContentTypesNamespace;
            if (contentTypes.Root?.Name != ns + "Types") {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package content-type manifest namespace is not supported by the native path.");
            }

            string expectedPartName = "/" + workbookPartName.TrimStart('/');
            string?[] matches = contentTypes.Root
                .Elements(ns + "Override")
                .Where(element => string.Equals(
                    NormalizeContentTypePartName((string?)element.Attribute("PartName")),
                    expectedPartName,
                    StringComparison.OrdinalIgnoreCase))
                .Select(element => (string?)element.Attribute("ContentType"))
                .Take(2)
                .ToArray();
            if (matches.Length != 1
                || string.IsNullOrWhiteSpace(matches[0])
                || !SupportedWorkbookContentTypes.Contains(matches[0]!)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The workbook content type is not supported by the native path.");
            }
        }

        private IReadOnlyDictionary<string, PackageRelationship> ReadRelationships(
            string sourcePartName) {
            string relationshipPartName = GetRelationshipPartName(sourcePartName);
            XDocument relationships = ReadXmlPart(
                relationshipPartName,
                MaximumMetadataPartBytes);
            XNamespace ns = PackageRelationshipsNamespace;
            if (relationships.Root?.Name != ns + "Relationships") {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The workbook relationship namespace is not supported by the native path.");
            }

            var result = new Dictionary<string, PackageRelationship>(StringComparer.Ordinal);
            foreach (XElement element in relationships.Root.Elements(ns + "Relationship")) {
                _options.CancellationToken.ThrowIfCancellationRequested();
                string? id = (string?)element.Attribute("Id");
                string? type = (string?)element.Attribute("Type");
                string? target = (string?)element.Attribute("Target");
                if (string.IsNullOrWhiteSpace(id)
                    || string.IsNullOrWhiteSpace(type)
                    || string.IsNullOrWhiteSpace(target)
                    || result.ContainsKey(id!)) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The workbook relationships require the Open XML SDK fallback path.");
                }

                result.Add(
                    id!,
                    new PackageRelationship(
                        type!,
                        target!,
                        string.Equals(
                            (string?)element.Attribute("TargetMode"),
                            "External",
                            StringComparison.OrdinalIgnoreCase)));
            }

            return result;
        }

        private (XlsxTabularSheet[] Sheets, ExcelDateSystem DateSystem) ReadWorkbook(
            string workbookPartName,
            IReadOnlyDictionary<string, PackageRelationship> relationships) {
            XDocument workbook = ReadXmlPart(workbookPartName, MaximumMetadataPartBytes);
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

            XNamespace transitionalRelationships = TransitionalOfficeRelationshipsNamespace;
            XNamespace strictRelationships = StrictOfficeRelationshipsNamespace;
            var sheets = new List<XlsxTabularSheet>();
            foreach (XElement sheet in sheetsElement.Elements(spreadsheet + "sheet")) {
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
                        continue;
                    }

                    throw new XlsxTabularFastPathNotSupportedException(
                        "A workbook sheet relationship requires the Open XML SDK fallback path.");
                }

                string partName = ResolveTarget(workbookPartName, relationship.Target);
                if (!_parts.ContainsPart(partName)) {
                    throw new InvalidDataException(
                        $"The OpenXML worksheet '{name}' references missing relationship '{relationshipId}'.");
                }
                sheets.Add(new XlsxTabularSheet(name!, partName));
            }

            return (sheets.ToArray(), dateSystem);
        }

        private string? ResolveOptionalPart(
            string workbookPartName,
            IReadOnlyDictionary<string, PackageRelationship> relationships,
            string relationshipSuffix,
            string relationshipName) {
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
                return XDocument.Load(reader, LoadOptions.None);
            } catch (XmlException exception) {
                throw new XlsxTabularFastPathNotSupportedException(
                    $"Package part '{partName}' requires the Open XML SDK fallback path.",
                    exception);
            }
        }

        private static string ResolveTarget(string sourcePartName, string target) {
            if (string.IsNullOrWhiteSpace(target)
                || target.IndexOf('\\') >= 0
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
            foreach (string segment in combined.Split('/')) {
                if (segment.Length == 0 || segment == ".") {
                    continue;
                }
                if (segment == "..") {
                    if (segments.Count == 0) {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "A package relationship target escapes the package root.");
                    }
                    segments.RemoveAt(segments.Count - 1);
                    continue;
                }
                if (segment.IndexOf('%') >= 0) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "An encoded package relationship target requires the Open XML SDK fallback path.");
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
            if (string.IsNullOrWhiteSpace(partName) || partName!.IndexOf('\\') >= 0) {
                return string.Empty;
            }

            return "/" + partName.TrimStart('/');
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
                _parts.Dispose();
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
