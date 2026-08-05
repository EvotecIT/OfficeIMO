using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.Utilities;
using System.IO;
using System.IO.Packaging;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Reader for an Excel workbook (read-only). Provides access to sheet readers and basic metadata.
    /// </summary>
    internal sealed partial class ExcelDocumentReader : IDisposable {
        private const string OfficeDocumentRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string StrictOfficeDocumentRelationshipNamespace = "http://purl.oclc.org/ooxml/officeDocument/relationships";
        private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string StrictSpreadsheetNamespace = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        private static readonly XmlReaderSettings WorkbookXmlReaderSettings = CreateWorkbookXmlReaderSettings();

        private readonly SpreadsheetDocument _doc;
        private readonly bool _owns;
        private readonly bool _canStreamWorksheetParts;
        private readonly ExcelReadOptions _opt;
        private readonly ExcelDateSystem _dateSystem;
        private readonly SharedStringCache _sst;
        private readonly StylesCacheProvider _styles;
        private readonly Package? _ownedPackage;
        private readonly Stream? _ownedStream;
        private readonly IDisposable? _ownedResource;
        private readonly OpenXmlPackagePartBufferReader? _partBufferReader;

        private ExcelDocumentReader(
            SpreadsheetDocument doc,
            ExcelReadOptions opt,
            bool owns,
            Package? ownedPackage = null,
            Stream? ownedStream = null,
            OpenXmlPackagePartBufferReader? partBufferReader = null,
            IDisposable? ownedResource = null) {
            _doc = doc;
            _owns = owns;
            _canStreamWorksheetParts = owns || doc.FileOpenAccess == FileAccess.Read;
            _opt = opt ?? new ExcelReadOptions();
            _opt.CancellationToken.ThrowIfCancellationRequested();
            _ownedPackage = ownedPackage;
            _ownedStream = ownedStream;
            _ownedResource = ownedResource;
            _partBufferReader = partBufferReader;
            _dateSystem = GetWorkbookDateSystem(doc);
            _sst = SharedStringCache.Build(doc, _opt);
            _styles = new StylesCacheProvider(doc);
            _opt.CancellationToken.ThrowIfCancellationRequested();
        }

        /// <summary>
        /// Opens an Excel file for read-only access.
        /// </summary>
        public static ExcelDocumentReader Open(string path, ExcelReadOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            if (!File.Exists(path)) {
                throw new FileNotFoundException($"File '{path}' doesn't exist.", path);
            }

            var effectiveOptions = options ?? new ExcelReadOptions();
            SharedReadOnlyFileSnapshot? snapshot = null;
            Stream? documentStream = null;
            SpreadsheetDocument? document = null;
            OpenXmlPackagePartBufferReader? partBufferReader = null;
            bool packageOpenAttempted = false;
            try {
                snapshot = SharedReadOnlyFileSnapshot.Open(path);
                if (snapshot.Length > effectiveOptions.MaxInputBytes) {
                    throw new InvalidDataException(
                        $"Workbook input contains {snapshot.Length} bytes, exceeding the configured limit of {effectiveOptions.MaxInputBytes} bytes.");
                }

                documentStream = snapshot.CreateView();
                packageOpenAttempted = true;
                document = SpreadsheetDocument.Open(documentStream, isEditable: false);
                partBufferReader = OpenXmlPackagePartBufferReader.TryOpen(snapshot.CreateView(bufferSize: 1));
                ExcelDocumentReader reader = new ExcelDocumentReader(
                    document,
                    effectiveOptions,
                    owns: true,
                    ownedStream: documentStream,
                    partBufferReader: partBufferReader,
                    ownedResource: snapshot);
                snapshot = null;
                documentStream = null;
                partBufferReader = null;
                return reader;
            } catch (Exception ex) {
                partBufferReader?.Dispose();
                document?.Dispose();
                documentStream?.Dispose();
                if (!packageOpenAttempted || !IsRecoverableOpenException(ex)) {
                    snapshot?.Dispose();
                    throw;
                }

                byte[] bytes;
                try {
                    using (Stream stream = snapshot!.CreateView()) {
                        bytes = OfficeStreamReader.ReadAllBytes(
                            stream,
                            effectiveOptions.CancellationToken,
                            effectiveOptions.MaxInputBytes);
                    }
                } finally {
                    snapshot?.Dispose();
                }

                return OpenFromBytes(
                    bytes,
                    effectiveOptions,
                    normalizeContentTypes: true,
                    contextMessage: $"Failed to open '{path}' after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.");
            }
        }

        /// <summary>
        /// Opens an Excel workbook from the provided stream for read-only access.
        /// </summary>
        public static ExcelDocumentReader Open(Stream stream, ExcelReadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            var effectiveOptions = options ?? new ExcelReadOptions();
            return OpenFromBytes(
                OfficeStreamReader.ReadAllBytes(
                    stream,
                    effectiveOptions.CancellationToken,
                    effectiveOptions.MaxInputBytes),
                effectiveOptions,
                normalizeContentTypes: false,
                contextMessage: "Failed to open workbook stream after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.");
        }

        /// <summary>
        /// Opens an Excel workbook from an in-memory package for read-only access.
        /// The byte array is used directly; callers should not modify it while the reader is alive.
        /// </summary>
        public static ExcelDocumentReader Open(byte[] bytes, ExcelReadOptions? options = null) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));

            var effectiveOptions = options ?? new ExcelReadOptions();
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            if (bytes.LongLength > effectiveOptions.MaxInputBytes) {
                throw new InvalidDataException($"Workbook input contains {bytes.LongLength} bytes, exceeding the configured limit of {effectiveOptions.MaxInputBytes} bytes.");
            }

            return OpenFromBytes(
                bytes,
                effectiveOptions,
                normalizeContentTypes: false,
                contextMessage: "Failed to open workbook bytes after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.");
        }

        /// <summary>
        /// Wraps an already open SpreadsheetDocument without taking ownership.
        /// The returned reader must be disposed, but it will not close the underlying document.
        /// </summary>
        public static ExcelDocumentReader Wrap(SpreadsheetDocument document, ExcelReadOptions? options = null) {
            if (document is null) throw new ArgumentNullException(nameof(document));
            return new ExcelDocumentReader(document, options ?? new ExcelReadOptions(), owns: false);
        }

        /// <summary>
        /// Returns the list of sheet names in workbook order.
        /// </summary>
        public IReadOnlyList<string> GetSheetNames() {
            if (TryGetSheetNamesXmlFast(out var fastNames)) {
                return fastNames;
            }

            var wb = WorkbookRoot;
            var names = new List<string>();
            foreach (var sheet in wb.Sheets!.Elements<Sheet>()) {
                if (!TryGetWorksheetPart(sheet, out _)) {
                    continue;
                }

                names.Add(sheet.Name!.Value!);
            }

            return names;
        }

        /// <summary>
        /// Gets a reader for the specified worksheet name.
        /// </summary>
        public ExcelSheetReader GetSheet(string name) {
            if (TryGetSheetByNameXmlFast(name, out string? fastSheetName, out WorksheetPart? fastWorksheetPart)) {
                return new ExcelSheetReader(fastSheetName, fastWorksheetPart, _sst, _styles, _opt, _dateSystem, _canStreamWorksheetParts, _partBufferReader);
            }

            var wb = WorkbookRoot;
            Sheet? sheet = null;
            foreach (var candidate in wb.Sheets!.Elements<Sheet>()) {
                if (string.Equals(candidate.Name?.Value, name, StringComparison.OrdinalIgnoreCase)
                    && TryGetWorksheetPart(candidate, out _)) {
                    sheet = candidate;
                    break;
                }
            }

            if (sheet is null) throw new KeyNotFoundException($"Sheet '{name}' not found.");
            if (!TryGetWorksheetPart(sheet, out WorksheetPart? wsPart)) {
                throw new KeyNotFoundException($"Sheet '{name}' is not a worksheet.");
            }

            return new ExcelSheetReader(sheet.Name!, wsPart!, _sst, _styles, _opt, _dateSystem, _canStreamWorksheetParts, _partBufferReader);
        }

        private bool TryGetWorksheetPart(Sheet sheet, out WorksheetPart? worksheetPart) {
            worksheetPart = null;
            if (sheet.Id?.Value == null) {
                return false;
            }

            try {
                worksheetPart = WorkbookPartRoot.GetPartById(sheet.Id.Value) as WorksheetPart;
                return worksheetPart != null;
            } catch (ArgumentOutOfRangeException) {
                return false;
            } catch (InvalidOperationException) {
                return false;
            }
        }

        private static ExcelDateSystem GetWorkbookDateSystem(SpreadsheetDocument document) {
            return document.WorkbookPart?.Workbook?.GetFirstChild<WorkbookProperties>()?.Date1904?.Value == true
                ? ExcelDateSystem.NineteenFour
                : ExcelDateSystem.NineteenHundred;
        }

        private bool TryGetSheetByNameXmlFast(string name, out string sheetName, out WorksheetPart worksheetPart) {
            sheetName = string.Empty;
            worksheetPart = null!;

            if (string.IsNullOrEmpty(name) || _doc.FileOpenAccess != FileAccess.Read) {
                return false;
            }

            try {
                using var stream = WorkbookPartRoot.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, WorkbookXmlReaderSettings);
                var state = new WorkbookSheetXmlState();
                string? resolvedName = null;
                WorksheetPart? resolvedWorksheetPart = null;
                while (reader.Read()) {
                    if (!IsWorkbookSheetElement(reader, state)) {
                        continue;
                    }

                    string? candidateName = reader.GetAttribute("name");
                    if (!string.Equals(candidateName, name, StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    string? relationshipId = GetRelationshipIdAttribute(reader);
                    if (string.IsNullOrEmpty(relationshipId)) {
                        return false;
                    }

                    if (WorkbookPartRoot.GetPartById(relationshipId!) is not WorksheetPart resolvedPart) {
                        return false;
                    }

                    resolvedName ??= candidateName;
                    resolvedWorksheetPart ??= resolvedPart;
                }

                if (state.IsInvalid || resolvedName == null || resolvedWorksheetPart == null) {
                    return false;
                }
                sheetName = resolvedName;
                worksheetPart = resolvedWorksheetPart;
                return true;
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            } catch (OpenXmlPackageException) {
                return false;
            } catch (InvalidOperationException) {
                return false;
            }
        }

        private static string? GetRelationshipIdAttribute(XmlReader reader) {
            string? relationshipId = reader.GetAttribute("id", OfficeDocumentRelationshipNamespace)
                ?? reader.GetAttribute("id", StrictOfficeDocumentRelationshipNamespace);
            if (!string.IsNullOrEmpty(relationshipId)) {
                return relationshipId;
            }

            if (!reader.HasAttributes || !reader.MoveToFirstAttribute()) {
                return null;
            }

            try {
                do {
                    if (reader.LocalName == "id"
                        && (reader.NamespaceURI == OfficeDocumentRelationshipNamespace
                            || reader.NamespaceURI == StrictOfficeDocumentRelationshipNamespace)) {
                        return reader.Value;
                    }
                } while (reader.MoveToNextAttribute());
            } finally {
                reader.MoveToElement();
            }

            return null;
        }

        private bool TryGetSheetByIndexXmlFast(int index, out string sheetName, out WorksheetPart worksheetPart) {
            sheetName = string.Empty;
            worksheetPart = null!;

            if (index < 1 || _doc.FileOpenAccess != FileAccess.Read) {
                return false;
            }

            try {
                using var stream = WorkbookPartRoot.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, WorkbookXmlReaderSettings);
                int currentIndex = 0;
                var state = new WorkbookSheetXmlState();
                string? resolvedName = null;
                WorksheetPart? resolvedWorksheetPart = null;
                while (reader.Read()) {
                    if (!IsWorkbookSheetElement(reader, state)) {
                        continue;
                    }

                    string? candidateName = reader.GetAttribute("name");
                    string? relationshipId = GetRelationshipIdAttribute(reader);
                    if (string.IsNullOrEmpty(candidateName) || string.IsNullOrEmpty(relationshipId)) {
                        continue;
                    }

                    if (WorkbookPartRoot.GetPartById(relationshipId!) is not WorksheetPart resolvedPart) {
                        continue;
                    }

                    currentIndex++;
                    if (currentIndex != index) {
                        continue;
                    }

                    resolvedName ??= candidateName;
                    resolvedWorksheetPart ??= resolvedPart;
                }

                if (state.IsInvalid || resolvedName == null || resolvedWorksheetPart == null) {
                    return false;
                }
                sheetName = resolvedName;
                worksheetPart = resolvedWorksheetPart;
                return true;
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            } catch (OpenXmlPackageException) {
                return false;
            } catch (InvalidOperationException) {
                return false;
            }
        }

        private bool TryGetSheetNamesXmlFast(out List<string> names) {
            names = [];

            if (_doc.FileOpenAccess != FileAccess.Read) {
                return false;
            }

            try {
                using var stream = WorkbookPartRoot.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, WorkbookXmlReaderSettings);
                var state = new WorkbookSheetXmlState();
                while (reader.Read()) {
                    if (!IsWorkbookSheetElement(reader, state)) {
                        continue;
                    }

                    string? sheetName = reader.GetAttribute("name");
                    if (string.IsNullOrEmpty(sheetName)) {
                        return false;
                    }

                    string? relationshipId = GetRelationshipIdAttribute(reader);
                    if (string.IsNullOrEmpty(relationshipId)
                        || WorkbookPartRoot.GetPartById(relationshipId!) is not WorksheetPart) {
                        continue;
                    }

                    names.Add(sheetName!);
                }

                return !state.IsInvalid && names.Count > 0;
            } catch (XmlException) {
                names = [];
                return false;
            } catch (IOException) {
                names = [];
                return false;
            } catch (UnauthorizedAccessException) {
                names = [];
                return false;
            } catch (ObjectDisposedException) {
                names = [];
                return false;
            } catch (OpenXmlPackageException) {
                names = [];
                return false;
            } catch (InvalidOperationException) {
                names = [];
                return false;
            }
        }

        private bool TryGetSheetCountXmlFast(out int count) {
            count = 0;

            if (_doc.FileOpenAccess != FileAccess.Read) {
                return false;
            }

            try {
                using var stream = WorkbookPartRoot.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, WorkbookXmlReaderSettings);
                var state = new WorkbookSheetXmlState();
                while (reader.Read()) {
                    if (!IsWorkbookSheetElement(reader, state)) {
                        continue;
                    }

                    string? relationshipId = GetRelationshipIdAttribute(reader);
                    if (!string.IsNullOrEmpty(relationshipId)
                        && WorkbookPartRoot.GetPartById(relationshipId!) is WorksheetPart) {
                        count++;
                    }
                }

                return !state.IsInvalid && count > 0;
            } catch (XmlException) {
                count = 0;
                return false;
            } catch (IOException) {
                count = 0;
                return false;
            } catch (UnauthorizedAccessException) {
                count = 0;
                return false;
            } catch (ObjectDisposedException) {
                count = 0;
                return false;
            } catch (OpenXmlPackageException) {
                count = 0;
                return false;
            } catch (InvalidOperationException) {
                count = 0;
                return false;
            }
        }

        private static bool IsWorkbookSheetElement(XmlReader reader, WorkbookSheetXmlState state) {
            bool isSpreadsheetNamespace = reader.NamespaceURI == SpreadsheetNamespace
                || reader.NamespaceURI == StrictSpreadsheetNamespace;

            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 0) {
                if (!isSpreadsheetNamespace || reader.LocalName != "workbook") {
                    state.IsInvalid = true;
                } else {
                    state.WorkbookDepth = reader.Depth;
                    state.WorkbookNamespace = reader.NamespaceURI;
                }
                return false;
            }

            if (reader.NodeType == XmlNodeType.EndElement
                && state.SheetsDepth == reader.Depth
                && reader.LocalName == "sheets"
                && reader.NamespaceURI == state.WorkbookNamespace) {
                state.SheetsDepth = -1;
                return false;
            }

            if (reader.NodeType != XmlNodeType.Element) {
                return false;
            }

            if (isSpreadsheetNamespace && reader.LocalName == "sheets") {
                if (state.WorkbookDepth < 0
                    || state.SheetsContainerSeen
                    || reader.Depth != state.WorkbookDepth + 1
                    || reader.NamespaceURI != state.WorkbookNamespace) {
                    state.IsInvalid = true;
                    return false;
                }
                state.SheetsContainerSeen = true;
                state.SheetsDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                return false;
            }

            return state.SheetsDepth >= 0
                && reader.Depth == state.SheetsDepth + 1
                && reader.LocalName == "sheet"
                && reader.NamespaceURI == state.WorkbookNamespace;
        }

        private sealed class WorkbookSheetXmlState {
            internal int WorkbookDepth { get; set; } = -1;
            internal int SheetsDepth { get; set; } = -1;
            internal string? WorkbookNamespace { get; set; }
            internal bool SheetsContainerSeen { get; set; }
            internal bool IsInvalid { get; set; }
        }

        private static XmlReaderSettings CreateWorkbookXmlReaderSettings() {
            return new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true,
                CloseInput = false
            };
        }

        /// <summary>
        /// Disposes the underlying OpenXML document.
        /// </summary>
        public void Dispose() {
            if (_owns) {
                try {
                    _doc.Dispose();
                } finally {
                    try {
                        _ownedPackage?.Close();
                    } finally {
                        try {
                            _ownedStream?.Dispose();
                        } finally {
                            try {
                                _partBufferReader?.Dispose();
                            } finally {
                                _ownedResource?.Dispose();
                            }
                        }
                    }
                }
            }
        }

        private static ExcelDocumentReader OpenFromBytes(byte[] bytes, ExcelReadOptions? options, bool normalizeContentTypes, string contextMessage) {
            var effectiveOptions = options ?? new ExcelReadOptions();
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            if (bytes.LongLength > effectiveOptions.MaxInputBytes) {
                throw new InvalidDataException($"Workbook input contains {bytes.LongLength} bytes, exceeding the configured limit of {effectiveOptions.MaxInputBytes} bytes.");
            }

            MemoryStream? packageStream = null;
            Package? package = null;
            SpreadsheetDocument? document = null;
            OpenXmlPackagePartBufferReader? partBufferReader = null;
            try {
                if (normalizeContentTypes) {
                    packageStream = new MemoryStream(bytes.Length + 4096);
                    packageStream.Write(bytes, 0, bytes.Length);
                    packageStream.Position = 0;
                    ExcelPackageUtilities.NormalizeContentTypes(packageStream, leaveOpen: true);
                    effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
                    packageStream.Position = 0;
                } else {
                    packageStream = new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: false);
                }

                package = Package.Open(packageStream, FileMode.Open, FileAccess.Read);
                effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
                document = SpreadsheetDocument.Open(package);
                partBufferReader = normalizeContentTypes
                    ? null
                    : OpenXmlPackagePartBufferReader.TryOpen(bytes);
                ExcelDocumentReader reader = new ExcelDocumentReader(
                    document,
                    effectiveOptions,
                    owns: true,
                    package,
                    packageStream,
                    partBufferReader);
                partBufferReader = null;
                return reader;
            } catch (Exception ex) when (!normalizeContentTypes && IsRecoverableOpenException(ex)) {
                partBufferReader?.Dispose();
                document?.Dispose();
                package?.Close();
                packageStream?.Dispose();
                return OpenFromBytes(bytes, effectiveOptions, normalizeContentTypes: true, contextMessage);
            } catch (Exception ex) when (IsRecoverableOpenException(ex)) {
                partBufferReader?.Dispose();
                document?.Dispose();
                package?.Close();
                packageStream?.Dispose();
                throw new IOException($"{contextMessage} See inner exception for details.", ex);
            } catch {
                partBufferReader?.Dispose();
                document?.Dispose();
                package?.Close();
                packageStream?.Dispose();
                throw;
            }
        }

        private static bool IsRecoverableOpenException(Exception ex) {
            return ex is InvalidDataException || ex is OpenXmlPackageException || ex is XmlException;
        }

    }
}
