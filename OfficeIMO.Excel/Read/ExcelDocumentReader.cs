using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using OfficeIMO.Excel.Utilities;
using System.IO;
using System.IO.Packaging;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Reader for an Excel workbook (read-only). Provides access to sheet readers and basic metadata.
    /// </summary>
    public sealed partial class ExcelDocumentReader : IDisposable {
        private const string OfficeDocumentRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string StrictOfficeDocumentRelationshipNamespace = "http://purl.oclc.org/ooxml/officeDocument/relationships";
        private static readonly XmlReaderSettings WorkbookXmlReaderSettings = CreateWorkbookXmlReaderSettings();

        private readonly SpreadsheetDocument _doc;
        private readonly bool _owns;
        private readonly ExcelReadOptions _opt;
        private readonly SharedStringCache _sst;
        private readonly StylesCacheProvider _styles;
        private readonly Package? _ownedPackage;
        private readonly Stream? _ownedStream;

        private ExcelDocumentReader(SpreadsheetDocument doc, ExcelReadOptions opt, bool owns, Package? ownedPackage = null, Stream? ownedStream = null) {
            _doc = doc;
            _owns = owns;
            _opt = opt ?? new ExcelReadOptions();
            _ownedPackage = ownedPackage;
            _ownedStream = ownedStream;
            _sst = SharedStringCache.Build(doc, _opt);
            _styles = new StylesCacheProvider(doc);
        }

        /// <summary>
        /// Opens an Excel file for read-only access.
        /// </summary>
        public static ExcelDocumentReader Open(string path, ExcelReadOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            if (!File.Exists(path)) {
                throw new FileNotFoundException($"File '{path}' doesn't exist.", path);
            }

            return OpenFromBytes(
                File.ReadAllBytes(path),
                options,
                normalizeContentTypes: false,
                contextMessage: $"Failed to open '{path}' after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.");
        }

        /// <summary>
        /// Opens an Excel workbook from the provided stream for read-only access.
        /// </summary>
        public static ExcelDocumentReader Open(Stream stream, ExcelReadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            return OpenFromBytes(
                ReadAllBytes(stream),
                options,
                normalizeContentTypes: false,
                contextMessage: "Failed to open workbook stream after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.");
        }

        /// <summary>
        /// Opens an Excel workbook from an in-memory package for read-only access.
        /// The byte array is used directly; callers should not modify it while the reader is alive.
        /// </summary>
        public static ExcelDocumentReader Open(byte[] bytes, ExcelReadOptions? options = null) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));

            return OpenFromBytes(
                bytes,
                options,
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
                names.Add(sheet.Name!.Value!);
            }

            return names;
        }

        /// <summary>
        /// Gets a reader for the specified worksheet name.
        /// </summary>
        public ExcelSheetReader GetSheet(string name) {
            if (TryGetSheetByNameXmlFast(name, out string? fastSheetName, out WorksheetPart? fastWorksheetPart)) {
                return new ExcelSheetReader(fastSheetName, fastWorksheetPart, _sst, _styles, _opt, _owns);
            }

            var wb = WorkbookRoot;
            Sheet? sheet = null;
            foreach (var candidate in wb.Sheets!.Elements<Sheet>()) {
                if (string.Equals(candidate.Name?.Value, name, StringComparison.OrdinalIgnoreCase)) {
                    sheet = candidate;
                    break;
                }
            }

            if (sheet is null) throw new KeyNotFoundException($"Sheet '{name}' not found.");
            var wsPart = (WorksheetPart)WorkbookPartRoot.GetPartById(sheet.Id!);
            return new ExcelSheetReader(sheet.Name!, wsPart, _sst, _styles, _opt, _owns);
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
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "sheet") {
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

                    sheetName = candidateName!;
                    worksheetPart = resolvedPart;
                    return true;
                }

                return false;
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
                ?? reader.GetAttribute("id", StrictOfficeDocumentRelationshipNamespace)
                ?? reader.GetAttribute("r:id");
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
                            || reader.NamespaceURI == StrictOfficeDocumentRelationshipNamespace
                            || reader.Prefix == "r")) {
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
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "sheet") {
                        continue;
                    }

                    currentIndex++;
                    if (currentIndex != index) {
                        continue;
                    }

                    string? candidateName = reader.GetAttribute("name");
                    string? relationshipId = GetRelationshipIdAttribute(reader);
                    if (string.IsNullOrEmpty(candidateName) || string.IsNullOrEmpty(relationshipId)) {
                        return false;
                    }

                    if (WorkbookPartRoot.GetPartById(relationshipId!) is not WorksheetPart resolvedPart) {
                        return false;
                    }

                    sheetName = candidateName!;
                    worksheetPart = resolvedPart;
                    return true;
                }

                return false;
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
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "sheet") {
                        continue;
                    }

                    string? sheetName = reader.GetAttribute("name");
                    if (string.IsNullOrEmpty(sheetName)) {
                        return false;
                    }

                    names.Add(sheetName!);
                }

                return names.Count > 0;
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
                while (reader.Read()) {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheet") {
                        count++;
                    }
                }

                return count > 0;
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
                _doc.Dispose();
                _ownedPackage?.Close();
                _ownedStream?.Dispose();
            }
        }

        private static ExcelDocumentReader OpenFromBytes(byte[] bytes, ExcelReadOptions? options, bool normalizeContentTypes, string contextMessage) {
            MemoryStream? packageStream = null;
            Package? package = null;
            try {
                if (normalizeContentTypes) {
                    packageStream = new MemoryStream(bytes.Length + 4096);
                    packageStream.Write(bytes, 0, bytes.Length);
                    packageStream.Position = 0;
                    ExcelPackageUtilities.NormalizeContentTypes(packageStream, leaveOpen: true);
                    packageStream.Position = 0;
                } else {
                    packageStream = new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: false);
                }

                package = Package.Open(packageStream, FileMode.Open, FileAccess.Read);
                var doc = SpreadsheetDocument.Open(package);
                return new ExcelDocumentReader(doc, options ?? new ExcelReadOptions(), owns: true, package, packageStream);
            } catch (Exception ex) when (!normalizeContentTypes && IsRecoverableOpenException(ex)) {
                package?.Close();
                packageStream?.Dispose();
                return OpenFromBytes(bytes, options, normalizeContentTypes: true, contextMessage);
            } catch (Exception ex) when (IsRecoverableOpenException(ex)) {
                package?.Close();
                packageStream?.Dispose();
                throw new IOException($"{contextMessage} See inner exception for details.", ex);
            } catch {
                package?.Close();
                packageStream?.Dispose();
                throw;
            }
        }

        private static bool IsRecoverableOpenException(Exception ex) {
            return ex is InvalidDataException || ex is OpenXmlPackageException || ex is XmlException;
        }

        private static byte[] ReadAllBytes(Stream stream) {
            if (stream is MemoryStream memoryStream) {
                return memoryStream.ToArray();
            }

            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
                long length = stream.Length;
                if (length > int.MaxValue) {
                    throw new IOException("Workbook stream is too large to read into memory.");
                }

                var bytes = new byte[(int)length];
                int offset = 0;
                while (offset < bytes.Length) {
                    int read = stream.Read(bytes, offset, bytes.Length - offset);
                    if (read == 0) {
                        break;
                    }

                    offset += read;
                }

                if (offset == bytes.Length) {
                    return bytes;
                }

                Array.Resize(ref bytes, offset);
                return bytes;
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }
    }
}
