using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using OfficeIMO.Excel.Utilities;
using System.IO;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Reader for an Excel workbook (read-only). Provides access to sheet readers and basic metadata.
    /// </summary>
    public sealed partial class ExcelDocumentReader : IDisposable {
        private readonly SpreadsheetDocument _doc;
        private readonly bool _owns;
        private readonly ExcelReadOptions _opt;
        private readonly SharedStringCache _sst;
        private readonly StylesCache _styles;

        private ExcelDocumentReader(SpreadsheetDocument doc, ExcelReadOptions opt, bool owns) {
            _doc = doc;
            _owns = owns;
            _opt = opt ?? new ExcelReadOptions();
            _sst = SharedStringCache.Build(doc);
            _styles = StylesCache.Build(doc);
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
                $"Failed to open '{path}' after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.");
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
                "Failed to open workbook stream after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.");
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
            var wb = WorkbookRoot;
            return wb.Sheets!.Elements<Sheet>().Select(s => s.Name!.Value!).ToList();
        }

        /// <summary>
        /// Gets a reader for the specified worksheet name.
        /// </summary>
        public ExcelSheetReader GetSheet(string name) {
            var wb = WorkbookRoot;
            var sheet = wb.Sheets!.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
            if (sheet is null) throw new KeyNotFoundException($"Sheet '{name}' not found.");
            var wsPart = (WorksheetPart)WorkbookPartRoot.GetPartById(sheet.Id!);
            return new ExcelSheetReader(sheet.Name!, wsPart, _sst, _styles, _opt);
        }

        /// <summary>
        /// Disposes the underlying OpenXML document.
        /// </summary>
        public void Dispose() {
            if (_owns)
                _doc.Dispose();
        }

        private static ExcelDocumentReader OpenFromBytes(byte[] bytes, ExcelReadOptions? options, string contextMessage) {
            MemoryStream? normalizedStream = null;
            try {
                normalizedStream = new MemoryStream(bytes.Length + 4096);
                normalizedStream.Write(bytes, 0, bytes.Length);
                normalizedStream.Position = 0;

                ExcelPackageUtilities.NormalizeContentTypes(normalizedStream, leaveOpen: true);
                normalizedStream.Position = 0;

                var doc = SpreadsheetDocument.Open(normalizedStream, false);
                return new ExcelDocumentReader(doc, options ?? new ExcelReadOptions(), owns: true);
            } catch (Exception ex) when (ex is InvalidDataException || ex is OpenXmlPackageException || ex is XmlException) {
                normalizedStream?.Dispose();
                throw new IOException($"{contextMessage} See inner exception for details.", ex);
            } catch {
                normalizedStream?.Dispose();
                throw;
            }
        }

        private static byte[] ReadAllBytes(Stream stream) {
            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }
    }
}
