using System.Diagnostics;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>A sheet entry of a workbook read without the Open XML SDK.</summary>
    internal sealed class ExcelDeferredSheet {
        internal ExcelDeferredSheet(string name, uint sheetId, string relationshipId, string partName, bool isWorksheet) {
            Name = name;
            SheetId = sheetId;
            RelationshipId = relationshipId;
            PartName = partName;
            IsWorksheet = isWorksheet;
        }

        internal string Name { get; }
        internal uint SheetId { get; }
        internal string RelationshipId { get; }
        internal string PartName { get; }
        internal bool IsWorksheet { get; }
    }

    /// <summary>What a load reads from a workbook without the Open XML SDK.</summary>
    internal sealed class ExcelDeferredWorkbook {
        internal ExcelDeferredWorkbook(string workbookPartName, List<ExcelDeferredSheet> sheets, bool date1904) {
            WorkbookPartName = workbookPartName;
            Sheets = sheets;
            Date1904 = date1904;
        }

        internal bool Date1904 { get; }

        internal string WorkbookPartName { get; }
        internal List<ExcelDeferredSheet> Sheets { get; }
    }

    /// <summary>
    /// A plain, unsigned Transitional workbook loads without opening the Open XML SDK package. The load reads the
    /// workbook's sheet list with the owned package reader, and the SDK opens over the same bytes on first use, exactly
    /// as the load would have opened it. Every such opening is recorded in <see cref="OpenXmlSdkFallbacks"/>.
    /// </summary>
    public partial class ExcelDocument {
        private static readonly HashSet<string> WorkbookMainContentTypes = new(StringComparer.OrdinalIgnoreCase) {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
        };
        private const string WorksheetRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "worksheet";
        private const string DigitalSignatureOriginRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
        private const string CustomPropertiesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "custom-properties";

        // Set while the SDK package has not been opened: what the load read, and how the SDK opens when it is needed.
        private ExcelDeferredWorkbook? _deferredWorkbook;
        private byte[]? _deferredPackageBytes;
        private OpenSettings? _deferredOpenSettings;
        private bool _deferredReadOnly;
        private readonly List<string> _openXmlSdkFallbacks = new();

        /// <summary>True while the workbook has not opened its Open XML SDK package.</summary>
        internal bool IsSdkDeferred => _deferredWorkbook != null;

        /// <summary>
        /// Why the workbook opened the Open XML SDK after a load that did not need it: one entry per opening, naming the
        /// member that needed it. Empty while the workbook works through OfficeIMO's own Open XML layer.
        /// </summary>
        public IReadOnlyList<string> OpenXmlSdkFallbacks => _openXmlSdkFallbacks;

        /// <summary>
        /// Opens the SDK package of a deferred load over the bytes the load read, with its settings, and binds the
        /// worksheets the document has handed out. Does nothing when the package is already open.
        /// </summary>
        internal void OpenDeferredSdk(string? reason) {
            ExcelDeferredWorkbook? deferred = _deferredWorkbook;
            if (deferred == null) {
                if (_disposed && _spreadSheetDocumentValue == null) throw new ObjectDisposedException(nameof(ExcelDocument));
                return;
            }
            byte[] bytes = _deferredPackageBytes!;
            Stream stream = OfficePackageLoadSource.CreatePackageStream(bytes, DocumentAccessMode.ReadWrite);
            SpreadsheetDocument document;
            try {
                document = _deferredReadOnly
                    ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlReadOnlyPackage.OpenSpreadsheet(stream, _deferredOpenSettings!)
                    : SpreadsheetDocument.Open(stream, true, _deferredOpenSettings!);
            } catch {
                stream.Dispose();
                throw;
            }

            _deferredWorkbook = null;
            _deferredPackageBytes = null;
            _spreadSheetDocumentValue = document;
            _workBookPartValue = GetWorkbookPartOrThrow(document);
            _packageCopyOnWriteStream = stream as OfficeCopyOnWriteStream;
            _packageOnCopyOnWriteStream = _packageCopyOnWriteStream != null ? document : null;
            ExcelChartAxisIdGenerator.ReadChartRelationships(document);
            BindDeferredSheets(deferred);
            _openXmlSdkFallbacks.Add(reason ?? DescribeSdkCaller());
            // Values written before the package opened go through the model now, in the order they were written.
            ReplayDeferredCellEdits();
        }

        internal void ThrowIfDisposed() {
            if (_disposed) throw new ObjectDisposedException(nameof(ExcelDocument));
        }

        // Diagnostics only: names the first member outside the document's package accessors that needed the SDK. A
        // trimmed or NativeAOT image may describe its frames with less detail.
        private static string DescribeSdkCaller() {
            try {
                string trace = new StackTrace(2, false).ToString();
                foreach (string rawLine in trace.Split('\n')) {
                    string line = rawLine.Trim();
                    if (line.StartsWith("at ", StringComparison.Ordinal)) line = line.Substring(3);
                    int arguments = line.IndexOf('(');
                    if (arguments > 0) line = line.Substring(0, arguments);
                    if (line.Length == 0 || line.Contains(".get__") || line.Contains(nameof(OpenDeferredSdk)) || line.Contains("BindDeferredSheet")
                        || line.Contains("<")) {
                        continue;
                    }
                    int dot = line.LastIndexOf('.', line.Length - 1);
                    int typeDot = dot > 0 ? line.LastIndexOf('.', dot - 1) : -1;
                    return typeDot >= 0 ? line.Substring(typeDot + 1) : line;
                }
            } catch (Exception exception) when (exception is not OutOfMemoryException) {
            }
            return "Open XML SDK needed";
        }

        /// <summary>Binds each handed-out sheet to its SDK element and part, and rebuilds the sheet cache around them.</summary>
        private void BindDeferredSheets(ExcelDeferredWorkbook deferred) {
            var elements = ReadSheetElements();
            var byRelationship = new Dictionary<string, Sheet>(StringComparer.Ordinal);
            foreach (Sheet element in elements) {
                if (element.Id?.Value is { } id && !byRelationship.ContainsKey(id)) byRelationship[id] = element;
            }

            lock (_sheetWrapperCacheLock) {
                for (int index = _sheetWrappers.Count - 1; index >= 0; index--) {
                    if (!_sheetWrappers[index].TryGetTarget(out ExcelSheet? wrapper)) {
                        _sheetWrappers.RemoveAt(index);
                        continue;
                    }
                    ExcelDeferredSheet? entry = wrapper.DeferredEntry;
                    if (entry == null || !byRelationship.TryGetValue(entry.RelationshipId, out Sheet? element)) continue;
                    if (_workBookPartValue.GetPartById(entry.RelationshipId) is WorksheetPart part) {
                        wrapper.BindToSdk(_spreadSheetDocumentValue, element, part);
                    }
                }
            }

            lock (_sheetCacheRebuildGate) {
                UpdateSheetIdCache(elements);
                _cachedSheetElements = elements.ConvertAll(element => new CachedSheetElement(element));
                if (_cachedSheets != null) {
                    // The cached wrappers are the ones callers hold; they stay the document's sheets.
                    _cachedSheets = _cachedSheets.Where(sheet => sheet.DeferredEntry == null).ToList();
                    if (_cachedSheets.Count != elements.Count(IsWorksheetElement)) _sheetCacheDirty = true;
                }
            }
            _ = deferred;
        }

        private bool IsWorksheetElement(Sheet sheet) =>
            sheet.Id?.Value is { } id && _workBookPartValue.GetPartById(id) is WorksheetPart;

        /// <summary>The sheets of a deferred workbook, as the SDK route would hand them out: worksheets only.</summary>
        private List<ExcelSheet> MaterializeDeferredSheets(ExcelDeferredWorkbook deferred) {
            var sheets = new List<ExcelSheet>(deferred.Sheets.Count);
            foreach (ExcelDeferredSheet entry in deferred.Sheets) {
                if (entry.IsWorksheet) sheets.Add(new ExcelSheet(this, entry));
            }
            return sheets;
        }

        private void UpdateDeferredSheetIdCache(ExcelDeferredWorkbook deferred) {
            id.Clear();
            id.Add(0);
            var seen = new HashSet<uint> { 0 };
            foreach (ExcelDeferredSheet entry in deferred.Sheets) {
                if (seen.Add(entry.SheetId)) id.Add(entry.SheetId);
            }
        }

        // Settings beyond the defaults change how the SDK reads parts, so those loads open it right away.
        private static bool IsPlainOpenSettings(OpenSettings settings) =>
            settings.MaxCharactersInPart == 0
            && settings.MarkupCompatibilityProcessSettings.ProcessMode == MarkupCompatibilityProcessMode.NoProcess;

        /// <summary>
        /// Reads the workbook of a plain, unsigned Transitional package whose every sheet entry reaches a part, or
        /// returns null for the load to open the SDK as before.
        /// </summary>
        internal static ExcelDeferredWorkbook? TryReadDeferredWorkbook(byte[] bytes, OpenSettings settings) {
            if (!IsPlainOpenSettings(settings)) return null;
            try {
                using var source = new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: true);
                using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
                foreach (OfficeOpcRelationship relationship in package.GetRelationships(null)) {
                    // Custom properties load through the SDK; a signed package keeps the SDK's signature handling.
                    if (string.Equals(relationship.Type, DigitalSignatureOriginRelationshipType, StringComparison.Ordinal)
                        || string.Equals(relationship.Type, CustomPropertiesRelationshipType, StringComparison.Ordinal)) {
                        return null;
                    }
                }
                string? main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
                if (main == null || !package.PartExists(main) || package.GetContentType(main) is not { } mainType
                    || !WorkbookMainContentTypes.Contains(mainType)) {
                    return null;
                }
                return ReadDeferredSheets(package, main);
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException
                or ArgumentException or FormatException or NotSupportedException or InvalidOperationException or OverflowException) {
                return null;
            }
        }

        // These reads leave markup compatibility as written, as the SDK's default NoProcess mode reads it.
        private static ExcelDeferredWorkbook? ReadDeferredSheets(OfficeOpcPackage package, string main) {
            var sheets = new List<ExcelDeferredSheet>();
            OfficeXmlPullReader reader = package.OpenXml(main, out OfficePooledBytes content, processMarkupCompatibility: false);
            using (content) {
                bool root = false, inSheets = false, sheetsSeen = false, propertiesSeen = false, date1904 = false;
                while (reader.Read()) {
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                    if (reader.Depth == 1) {
                        // Strict markup goes through the SDK, which maps it.
                        if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "workbook")
                            || reader.ElementUriId != OfficeXmlNamespaces.TransitionalUriId(OfficeXmlNamespaces.SpreadsheetML)) {
                            return null;
                        }
                        root = true;
                        continue;
                    }
                    if (reader.Depth == 2 && !propertiesSeen && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "workbookPr")) {
                        // The date system comes from the first workbook properties element, as the SDK finds it.
                        propertiesSeen = true;
                        if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "date1904", out string date1904Text)) {
                            if (date1904Text is "1" or "true") date1904 = true;
                            else if (date1904Text is not ("0" or "false")) return null;
                        }
                    }
                    if (reader.Depth == 2) {
                        // The SDK reads the first sheet list under the root.
                        inSheets = !sheetsSeen && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "sheets");
                        sheetsSeen |= inSheets;
                        continue;
                    }
                    if (reader.Depth != 3 || !inSheets || !reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "sheet")) continue;
                    if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "name", out string name)
                        || !reader.TryGetAttribute(OfficeXmlNamespaces.None, "sheetId", out string sheetIdText)
                        || !uint.TryParse(sheetIdText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out uint sheetId)
                        || !reader.TryGetAttribute(OfficeXmlNamespaces.Relationships, "id", out string id)) {
                        return null;
                    }
                    // The SDK resolves each entry's relationship when the sheets are listed and fails on one it cannot.
                    OfficeOpcRelationship? relationship = package.FindRelationship(main, id);
                    if (relationship == null || relationship.IsExternal || relationship.TargetPartName is not { } target
                        || !package.PartExists(target)) {
                        return null;
                    }
                    // The SDK types a part by the relationship that reaches it.
                    bool isWorksheet = string.Equals(relationship.Type, WorksheetRelationshipType, StringComparison.Ordinal);
                    sheets.Add(new ExcelDeferredSheet(name, sheetId, id, target, isWorksheet));
                }
                return root ? new ExcelDeferredWorkbook(main, sheets, date1904) : null;
            }
        }
    }
}
