using OfficeIMO.Shared;
using OfficeIMO.Word.LegacyDoc.Diagnostics;

namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Neutral legacy binary Word document model for the supported import subset.
    /// </summary>
    public sealed class LegacyDocDocument {
        private readonly List<LegacyDocImportDiagnostic> _diagnostics = new();
        private readonly List<string> _paragraphs = new();

        private LegacyDocDocument() {
        }

        /// <summary>Gets body text decoded from the Word piece table.</summary>
        public string Text { get; private set; } = string.Empty;

        /// <summary>Gets body paragraphs projected from Word paragraph marks.</summary>
        public IReadOnlyList<string> Paragraphs => _paragraphs;

        /// <summary>Gets diagnostics produced while reading the legacy document.</summary>
        public IReadOnlyList<LegacyDocImportDiagnostic> Diagnostics => _diagnostics;

        /// <summary>
        /// Loads a legacy DOC model from a file path.
        /// </summary>
        public static LegacyDocDocument Load(string path, LegacyDocImportOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            return Load(File.ReadAllBytes(path), options);
        }

        /// <summary>
        /// Loads a legacy DOC model from a stream.
        /// </summary>
        public static LegacyDocDocument Load(Stream stream, LegacyDocImportOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            using var buffer = new MemoryStream();
            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }
            stream.CopyTo(buffer);
            return Load(buffer.ToArray(), options);
        }

        /// <summary>
        /// Loads a legacy DOC model from compound document bytes.
        /// </summary>
        public static LegacyDocDocument Load(byte[] bytes, LegacyDocImportOptions? options = null) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            options ??= new LegacyDocImportOptions();

            var document = new LegacyDocDocument();
            if (!OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compoundFile, out string? compoundError)) {
                document.AddError("DOC-COMPOUND-INVALID", compoundError ?? "The OLE compound document could not be read.");
                return document;
            }

            document.LoadFromCompound(compoundFile!, options);
            return document;
        }

        /// <summary>
        /// Creates a compact report from the decoded model and diagnostics.
        /// </summary>
        public LegacyDocImportReport CreateImportReport() {
            return new LegacyDocImportReport(this);
        }

        private void LoadFromCompound(OfficeCompoundFile compoundFile, LegacyDocImportOptions options) {
            if (!compoundFile.Streams.TryGetValue("WordDocument", out byte[]? wordDocumentStream)) {
                AddError("DOC-WORDDOCUMENT-MISSING", "The compound document does not contain a WordDocument stream.");
                return;
            }

            if (wordDocumentStream.Length > options.MaxWordDocumentStreamBytes) {
                AddError("DOC-WORDDOCUMENT-TOO-LARGE", $"The WordDocument stream is {wordDocumentStream.Length} bytes, which exceeds the configured limit of {options.MaxWordDocumentStreamBytes} bytes.");
                return;
            }

            LegacyDocFib fib;
            if (!LegacyDocFib.TryRead(wordDocumentStream, out fib, out string? fibError)) {
                AddError("DOC-FIB-INVALID", fibError ?? "The WordDocument stream does not contain a supported File Information Block.");
                return;
            }

            if (fib.IsEncrypted) {
                AddError("DOC-ENCRYPTED", "Password-protected binary .doc files are detected but are not imported by the dependency-free reader.");
                return;
            }

            if (options.ReportUnsupportedFeatures) {
                AddKnownUnsupportedFeatureDiagnostics(compoundFile);
            }

            string tableStreamName = fib.UsesOneTableStream ? "1Table" : "0Table";
            if (!compoundFile.Streams.TryGetValue(tableStreamName, out byte[]? tableStream)) {
                string alternateName = fib.UsesOneTableStream ? "0Table" : "1Table";
                if (compoundFile.Streams.TryGetValue(alternateName, out tableStream)) {
                    AddWarning("DOC-TABLE-STREAM-FALLBACK", $"The FIB requested {tableStreamName}, but only {alternateName} was present. The available table stream was used.");
                } else {
                    AddError("DOC-TABLE-STREAM-MISSING", $"The compound document does not contain the {tableStreamName} table stream.");
                    return;
                }
            }

            if (!LegacyDocPieceTable.TryRead(wordDocumentStream, tableStream, fib, out string text, out string? textError)) {
                AddError("DOC-PIECE-TABLE-INVALID", textError ?? "The legacy DOC piece table could not be decoded.");
                return;
            }

            Text = NormalizeBodyText(text);
            foreach (string paragraph in SplitParagraphs(Text)) {
                _paragraphs.Add(paragraph);
            }
        }

        private void AddKnownUnsupportedFeatureDiagnostics(OfficeCompoundFile compoundFile) {
            if (compoundFile.Entries.Any(entry => entry.Path.StartsWith("_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase)
                || entry.Path.StartsWith("Macros", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("VBA", StringComparison.OrdinalIgnoreCase) >= 0)) {
                AddWarning("DOC-MACROS-PRESENT", "The legacy DOC contains VBA project storage. Macros are not projected into the OfficeIMO document.");
            }

            if (compoundFile.Entries.Any(entry => entry.Path.IndexOf("ObjectPool", StringComparison.OrdinalIgnoreCase) >= 0)) {
                AddWarning("DOC-OLE-OBJECTS-PRESENT", "The legacy DOC contains embedded OLE object storage. Embedded objects are not projected yet.");
            }
        }

        private static string NormalizeBodyText(string text) {
            var builder = new System.Text.StringBuilder(text.Length);
            foreach (char character in text) {
                switch (character) {
                    case '\0':
                    case '\a':
                        break;
                    case '\v':
                    case '\f':
                        builder.Append('\r');
                        break;
                    default:
                        if (!char.IsControl(character) || character == '\t' || character == '\r' || character == '\n') {
                            builder.Append(character);
                        }
                        break;
                }
            }

            return builder.ToString();
        }

        private static IEnumerable<string> SplitParagraphs(string text) {
            string[] paragraphs = text.Replace("\r\n", "\r").Replace('\n', '\r').Split(new[] { '\r' }, StringSplitOptions.None);
            int count = paragraphs.Length;
            if (count > 0 && paragraphs[count - 1].Length == 0) {
                count--;
            }

            for (int i = 0; i < count; i++) {
                yield return paragraphs[i];
            }
        }

        private void AddError(string code, string message) {
            _diagnostics.Add(new LegacyDocImportDiagnostic(code, LegacyDocDiagnosticSeverity.Error, message));
        }

        private void AddWarning(string code, string message) {
            _diagnostics.Add(new LegacyDocImportDiagnostic(code, LegacyDocDiagnosticSeverity.Warning, message));
        }
    }
}
