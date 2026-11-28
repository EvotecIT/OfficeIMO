using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Text;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>Workbook metadata and part access shared by read-only opens and editable workbook snapshots.</summary>
    internal sealed class ExcelReaderPackage : IDisposable {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
        private readonly Func<Stream> _openSource;
        private readonly OfficeOpcPackage _package;
        private readonly string _workbookName;
        private readonly object _relationsGate = new();

        private ExcelReaderPackage(Func<Stream> openSource, OfficeOpcPackage package, string workbookName,
            OfficeXmlNode workbook, OpenXmlPackagePartBufferReader parts) {
            _openSource = openSource;
            _package = package;
            _workbookName = workbookName;
            Workbook = workbook;
            Parts = parts;
        }

        internal OfficeXmlNode Workbook { get; }

        internal OpenXmlPackagePartBufferReader Parts { get; }

        internal static ExcelReaderPackage Open(Func<Stream> openSource, OpenXmlPackagePartBufferReader parts,
            ExcelReaderOptions? limits = null) {
            IReadOnlyCollection<string> names = parts.IndexedPartNames;
            OfficeOpcPackage package = OfficeOpcPackage.OpenFromParts(names, names.Count,
                name => parts.ReadMetadataPart(name, limits?.MaxMetadataPartBytes ?? int.MaxValue, limits?.CancellationToken ?? default),
                name => parts.TryGetLength(name, out long length) ? length : (long?)null, validateMetadata: true);
            try {
                OfficeOpcRelationship? workbookRelationship = null;
                foreach (OfficeOpcRelationship relationship in package.GetRelationships(null)) {
                    if (relationship.Type != OfficeOpcPackage.OfficeDocumentRelationshipType) continue;
                    if (workbookRelationship != null) throw new InvalidDataException("The package declares more than one workbook relationship.");
                    workbookRelationship = relationship;
                }
                string workbookName = workbookRelationship?.TargetPartName ?? throw new InvalidDataException("WorkbookPart is null.");
                string? contentType = package.GetContentType(workbookName);
                if (contentType == null || !ExcelReadWorkbookParts.WorkbookContentTypes.Contains(contentType)) {
                    throw new InvalidDataException($"The package workbook part '{workbookName}' has an unsupported content type.");
                }
                using OfficePooledBytes bytes = package.ReadPart(workbookName);
                byte[] markup = bytes.Array.AsSpan(0, bytes.Length).ToArray();
                OfficeXmlNode workbook = ParseRoot(markup, package.Namespaces, limits?.CancellationToken ?? default);
                if (!workbook.Is(Spreadsheet, "workbook")) throw new InvalidDataException("Workbook is null.");
                // Resolve the workbook relationship set before the immutable view is shared.
                foreach (OfficeOpcRelationship relationship in package.GetRelationships(workbookName)) {
                    string? expected = ExcelReadWorkbookParts.FixedNonWorksheetContentType(relationship.Type);
                    if (expected != null && relationship.TargetPartName is string target && package.PartExists(target)) {
                        package.RequireContentType(target, expected);
                    }
                }
                if (limits != null) {
                    int count = 0;
                    foreach (OfficeXmlNode sheet in workbook.Child(Spreadsheet, "sheets").Children(Spreadsheet, "sheet")) {
                        limits.CancellationToken.ThrowIfCancellationRequested();
                        if (++count > limits.MaxWorksheets) {
                            throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxWorksheets), limits.MaxWorksheets,
                                $"The OpenXML workbook contains more than the configured {limits.MaxWorksheets} worksheet definitions.");
                        }
                    }
                }

                return new ExcelReaderPackage(openSource, package, workbookName, workbook, parts);
            } catch {
                package.Dispose();
                throw;
            }
        }

        internal Stream OpenSource() => _openSource();

        internal (string? Type, string? Target, bool External) FindWorkbookRelationship(string id) {
            lock (_relationsGate) {
                OfficeOpcRelationship? relationship = _package.FindRelationship(_workbookName, id);
                return relationship == null ? default : (relationship.Type, relationship.TargetPartName, relationship.IsExternal);
            }
        }

        private string? FindWorkbookPart((string Relationship, string ContentType) kind) {
            lock (_relationsGate) {
                foreach (OfficeOpcRelationship relation in _package.GetRelationships(_workbookName)) {
                    if (relation.Type != kind.Relationship || relation.TargetPartName is not string target) continue;
                    _package.RequireContentType(target, kind.ContentType);
                    if (!_package.PartExists(target)) throw new InvalidDataException($"The workbook part '{target}' is missing.");
                    return target;
                }
            }
            return null;
        }

        internal void ValidateWorksheetContentType(string name) {
            string? contentType = _package.GetContentType(name);
            if (!string.Equals(contentType, ExcelReadWorkbookParts.WorksheetContentType, StringComparison.OrdinalIgnoreCase)) {
                var failure = new InvalidDataException($"The worksheet part '{name}' has the content type '{contentType}'.");
                throw new IOException("The package declares an unsupported worksheet content type.", failure);
            }
        }

        internal SharedStringCache CreateSharedStrings(ExcelReaderOptions options, bool keepUnmarkedWhitespace) {
            string? name = FindWorkbookPart(ExcelReadWorkbookParts.SharedStrings);
            if (name == null) return SharedStringCache.Empty(options);
            Stream Open() => Parts.OpenPart(name!, int.MaxValue, options.CancellationToken,
                maximumBufferedBytes: keepUnmarkedWhitespace ? SharedStringCache.MaximumIndexedPartBytes : 8192,
                acceptBufferedPrefix: keepUnmarkedWhitespace ? SharedStringCache.CanIndexSimpleAsciiPrefix : null);
            return SharedStringCache.BuildForPackage(Open, options, keepUnmarkedWhitespace);
        }

        internal StylesCacheProvider CreateStylesProvider() {
            string? name = FindWorkbookPart(ExcelReadWorkbookParts.Styles);
            return name == null ? new StylesCacheProvider(StylesCache.Empty())
                : new StylesCacheProvider(() => Parts.OpenPart(name!, int.MaxValue));
        }

        internal OfficeXmlNode ReadStyles(CancellationToken cancellationToken) {
            string? name = FindWorkbookPart(ExcelReadWorkbookParts.Styles);
            if (name == null) return default;
            using Stream stream = Parts.OpenPart(name!, int.MaxValue, cancellationToken);
            return ReadPartRoot(stream, cancellationToken);
        }

        internal RichValueErrorLookup ReadRichValueErrors(ExcelReaderOptions options) {
            string? metadata = FindWorkbookPart(ExcelReadWorkbookParts.Metadata);
            string? values = FindWorkbookPart(ExcelReadWorkbookParts.RichValues);
            string? structures = FindWorkbookPart(ExcelReadWorkbookParts.RichStructures);
            if (metadata == null || values == null || structures == null) return RichValueErrorLookup.Empty;
            int limit = Math.Min(options.MaxMetadataPartBytes, RichValueErrorLookup.MaximumPartBytes);
            foreach (string name in new[] { metadata, values, structures }) {
                if (Parts.TryGetLength(name, out long length)) RichValueErrorLookup.CheckPartLength(length, limit);
            }
            OfficeXmlNode Read(string name) {
                using Stream stream = Parts.OpenPart(name, limit, options.CancellationToken, limitName: nameof(ExcelReaderOptions.MaxMetadataPartBytes));
                return ReadPartRoot(stream, options.CancellationToken, limit);
            }
            return RichValueErrorLookup.FromRoots(Read(metadata), Read(values), Read(structures));
        }

        internal static OfficeXmlNode ReadPartRoot(Stream stream, CancellationToken cancellationToken, int? maximumBytes = null,
            bool validateMarkup = false) {
            byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, maximumBytes);
            return ParseRoot(bytes, new OfficeXmlNamespaces.Table(), cancellationToken, validateMarkup);
        }

        private static OfficeXmlNode ParseRoot(byte[] bytes, OfficeXmlNamespaces.Table namespaces, CancellationToken cancellationToken,
            bool validateMarkup = false) {
            bool transcode = DeclaresOtherEncoding(bytes, bytes.Length);
            // The general worksheet projection also handles input declined by the strict streaming reader.
            if (validateMarkup || transcode) ValidateMarkup(bytes, cancellationToken);
            if (transcode) {
                using var source = new MemoryStream(bytes, writable: false);
                using var xml = new XmlTextReader(source) { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                xml.MoveToContent();
                if (xml.Encoding is Encoding encoding && encoding.CodePage != Encoding.UTF8.CodePage) {
                    bytes = Encoding.Convert(encoding, Encoding.UTF8, bytes);
                }
            }
            try {
                OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, namespaces);
                return new OfficeXmlNode(tree, tree.Root);
            } catch (XmlException) {
                // Preserve the standard XML diagnostic when the owned parser encounters malformed markup.
                if (!validateMarkup && !transcode) ValidateMarkup(bytes, cancellationToken);
                throw;
            }
        }

        private static void ValidateMarkup(byte[] bytes, CancellationToken cancellationToken) {
            using var source = new MemoryStream(bytes, writable: false);
            using XmlReader validation = XmlReader.Create(source, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            while (validation.Read()) cancellationToken.ThrowIfCancellationRequested();
        }

        internal static bool DeclaresOtherEncoding(byte[] data, int length) {
            int start = length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF ? 3 : 0;
            ReadOnlySpan<byte> text = data.AsSpan(start, length - start);
            if (!text.StartsWith("<?xml"u8)) return false;
            int end = text.IndexOf("?>"u8);
            if (end < 0) return false;
            ReadOnlySpan<byte> declaration = text.Slice(0, end);
            int encoding = declaration.IndexOf("encoding"u8);
            if (encoding < 0) return false;
            ReadOnlySpan<byte> rest = SkipXmlWhitespace(declaration.Slice(encoding + 8));
            if (rest.IsEmpty || rest[0] != (byte)'=') return true;
            rest = SkipXmlWhitespace(rest.Slice(1));
            if (rest.IsEmpty || rest[0] is not ((byte)'"') and not ((byte)'\'')) return true;
            int close = rest.Slice(1).IndexOf(rest[0]);
            if (close < 0) return true;
            ReadOnlySpan<byte> name = rest.Slice(1, close);
            return !IsUtf8EncodingName(name);
        }

        private static ReadOnlySpan<byte> SkipXmlWhitespace(ReadOnlySpan<byte> value) {
            int index = 0;
            while (index < value.Length && value[index] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') index++;
            return value.Slice(index);
        }

        private static bool IsUtf8EncodingName(ReadOnlySpan<byte> name) {
            if (name.Length is not 4 and not 5) return false;
            Span<byte> lower = stackalloc byte[5];
            for (int index = 0; index < name.Length; index++) lower[index] = (byte)(name[index] | 0x20);
            return lower.Slice(0, name.Length).SequenceEqual("utf-8"u8) || lower.Slice(0, name.Length).SequenceEqual("utf8"u8);
        }

        internal IReadOnlyList<string> GetWorksheetTablePartNames(string worksheetName) {
            const string tableRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "table";
            lock (_relationsGate) {
                var names = new List<string>();
                foreach (OfficeOpcRelationship relation in _package.GetRelationships(worksheetName)) {
                    if (relation.Type == tableRelationship && relation.TargetPartName is string target && PartExists(target)) {
                        names.Add(target);
                    }
                }
                return names;
            }
        }

        internal bool PartExists(string name) => _package.PartExists(name);

        public void Dispose() {
            try {
                _package.Dispose();
            } finally {
                Parts.Dispose();
            }
        }
    }
}
