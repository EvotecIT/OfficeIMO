using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using System.Buffers;
using System.Data.Common;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Owns the minimal validated package state needed by the forward-only XLSX reader.
    /// Unsupported package shapes route back to <see cref="ExcelDocumentReader"/>.
    /// </summary>
    internal sealed partial class XlsxTabularWorkbook : IDisposable {
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
        // Excel and the Open XML SDK write macro sheets under the Office 2006 relationship namespace.
        private const string OfficeMacroSheetRelationshipType =
            "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
        private const string OfficeInternationalMacroSheetRelationshipType =
            "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
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
        private readonly ExcelReaderOptions _options;
        private readonly Lazy<RichValueErrorLookup> _richValueErrors;
        private readonly Dictionary<string, string> _contentTypeOverrides;
        private readonly Dictionary<string, string> _contentTypeDefaults;
        private readonly XlsxTabularSheet[] _sheets;
        private readonly string[] _tableNames;
        // Element names repeat across the metadata parts, so their pull readers share one table.
        private OfficeXmlNamespaces.Table? _metadataNamespaces;
        private OfficeXmlPullReader? _metadataReader;
        private bool _disposed;

        private XlsxTabularWorkbook(
            OpenXmlPackagePartBufferReader parts,
            OpenXmlPackagePartBufferReader? prefetchedParts,
            IDisposable? ownedResource,
            ExcelReaderOptions options,
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
            Dictionary<string, PackageRelationship> workbookRelationships =
                ReadRelationships(workbookPartName);
            _richValueErrors = new Lazy<RichValueErrorLookup>(() => ReadRichValueErrors(workbookPartName, workbookRelationships));
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

            // Prefer overlapping inflation with indexing when that bounded pipeline is eligible.
            // Starting a complete-part prefetch would otherwise serialize the indexing phase.
            XlsxTabularSheet? prefetchedSheet = options.EnableWorksheetPrefetch && _prefetchedParts?.AllowsPipelining != true
                ? ResolvePrefetchedSheet(_sheets, options)
                : null;
            _prefetchedSheetPartName = prefetchedSheet?.PartName;
            if (_prefetchedSheetPartName != null && _prefetchedParts != null) {
                _prefetchedParts.BeginPrefetch(
                    _prefetchedSheetPartName,
                    MaximumPrefetchedWorksheetBytes,
                    options.CancellationToken,
                    ExcelSheetReader.CanBufferUtf8WorksheetPrefix);
            }
            int maximumPartBytes = options.MaxInputBytes > int.MaxValue
                ? int.MaxValue
                : checked((int)options.MaxInputBytes);

            string partLimitName = options.MaxInputBytes > int.MaxValue ? "MaxPackagePartBytes" : nameof(ExcelReaderOptions.MaxInputBytes);
            string? sharedStringsPart = ResolveOptionalPart(
                workbookPartName,
                workbookRelationships,
                SharedStringsRelationshipSuffix,
                "shared-string table",
                SharedStringsContentType);
            _sharedStrings = sharedStringsPart == null
                ? SharedStringCache.Empty(options)
                : SharedStringCache.Build(
                    () => _parts.OpenPart(sharedStringsPart, maximumPartBytes, options.CancellationToken,
                        maximumBufferedBytes: SharedStringCache.MaximumIndexedPartBytes,
                        acceptBufferedPrefix: SharedStringCache.CanIndexSimpleAsciiPrefix,
                        limitName: partLimitName),
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
                    () => _parts.OpenPart(stylesPart, maximumPartBytes, options.CancellationToken,
                        limitName: partLimitName));
        }

        internal IReadOnlyList<string> TableNames => _tableNames;

        internal ExcelDateSystem DateSystem { get; }

        // Indexing may overlap foreground inflation only when metadata has an
        // independent archive. Explicit sequential policies and diagnostic callbacks
        // retain their existing execution and callback-thread behavior.
        private static bool CanPipelineWorksheetIndex(ExcelReaderOptions options) =>
            options.A1Range == null && options.Execution.AllowsBackgroundProcessing;

        internal static XlsxTabularWorkbook Open(string path, ExcelReaderOptions options) {
            return Open(path, options, metadataOnly: false);
        }

        internal static IReadOnlyList<string> ReadSheetNames(string path, ExcelReaderOptions options) {
            using XlsxTabularWorkbook workbook = Open(path, options, metadataOnly: true);
            options.CancellationToken.ThrowIfCancellationRequested();
            return workbook._tableNames.ToArray();
        }

        private static XlsxTabularWorkbook Open(
            string path,
            ExcelReaderOptions options,
            bool metadataOnly) {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);
            ArgumentNullException.ThrowIfNull(options);

            SharedReadOnlyFileSnapshot? snapshot = null;
            OpenXmlPackagePartBufferReader? parts = null;
            OpenXmlPackagePartBufferReader? prefetchedParts = null;
            try {
                snapshot = SharedReadOnlyFileSnapshot.Open(path);
                if (snapshot.Length > options.MaxInputBytes) {
                    throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxInputBytes), options.MaxInputBytes,
                        $"Workbook input contains {snapshot.Length} bytes, exceeding the configured limit of {options.MaxInputBytes} bytes.");
                }

                using (Stream directory = snapshot.CreateView(bufferSize: 1)) {
                    ExcelDocumentLoadRouting.ValidateZipPackage(directory, options.CancellationToken);
                }
                parts = OpenXmlPackagePartBufferReader.TryOpen(snapshot.CreateView(bufferSize: 1))
                    ?? throw new XlsxTabularFastPathNotSupportedException(
                        "The workbook is not a readable Open XML package.");
                if (!metadataOnly && (options.EnableWorksheetPrefetch || CanPipelineWorksheetIndex(options))) {
                    prefetchedParts = parts.TryOpenIndependent(snapshot.CreateView(bufferSize: 1),
                        allowPipelining: CanPipelineWorksheetIndex(options));
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

        internal static XlsxTabularWorkbook Open(byte[] bytes, ExcelReaderOptions options) {
            ArgumentNullException.ThrowIfNull(bytes);
            ArgumentNullException.ThrowIfNull(options);
            if (bytes.LongLength > options.MaxInputBytes) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxInputBytes), options.MaxInputBytes,
                    $"Workbook input contains {bytes.LongLength} bytes, exceeding the configured limit of {options.MaxInputBytes} bytes.");
            }

            OpenXmlPackagePartBufferReader parts = OpenXmlPackagePartBufferReader.TryOpen(bytes)
                ?? throw new XlsxTabularFastPathNotSupportedException(
                    "The workbook is not a readable Open XML package.");
            OpenXmlPackagePartBufferReader? prefetchedParts = null;
            try {
                prefetchedParts = options.EnableWorksheetPrefetch || CanPipelineWorksheetIndex(options)
                    ? parts.TryOpenIndependent(new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: false),
                        allowPipelining: CanPipelineWorksheetIndex(options))
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
                !_options.EnableWorksheetPrefetch
                    || _prefetchedParts?.AllowsPipelining == true
                    || string.Equals(sheet.PartName, _prefetchedSheetPartName, StringComparison.OrdinalIgnoreCase)
                        ? _prefetchedParts ?? _parts
                        : _parts,
                _richValueErrors);
            DbDataReader dataReader = string.IsNullOrWhiteSpace(_options.A1Range)
                ? (DbDataReader)reader.ReadUsedRangeAsDataReader(
                    hasHeaderRow,
                    schemaSampleRows: 0,
                    cancellationToken)
                : (DbDataReader)reader.ReadRangeAsDataReader(
                    _options.A1Range!,
                    hasHeaderRow,
                    chunkRows: Math.Min(ExcelSheetReader.DataReaderChunkRows, _options.MaxDataReaderChunkRows),
                    schemaSampleRows: 0,
                    ct: cancellationToken);

            return _options.InferSchema && _options.SchemaSampleRows > 0
                ? ExcelSchemaInferenceDataReader.Create(
                    dataReader,
                    _options.SchemaSampleRows,
                    _options.MaxDataReaderSchemaSampleRows,
                    _options.MaxDataReaderBufferedCells,
                    _options.Culture,
                    DateSystem,
                    cancellationToken)
                : dataReader;
        }

        private string ReadWorkbookPartName() {
            Dictionary<string, PackageRelationship> relationships =
                ReadRelationships(string.Empty);
            PackageRelationship? workbook = FindSingleOfficeRelationship(relationships, "/officeDocument", out bool multiple);
            if (workbook == null || multiple || workbook.IsExternal) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package does not contain one internal Office workbook relationship.");
            }

            string workbookPartName = ResolveTarget(string.Empty, workbook.Target);
            if (!_parts.ContainsPart(workbookPartName)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The package workbook relationship target is missing.");
            }

            return workbookPartName;
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

        private string? ResolveOptionalPart(
            string workbookPartName,
            Dictionary<string, PackageRelationship> relationships,
            string relationshipSuffix,
            string relationshipName,
            string expectedContentType) {
            PackageRelationship? match = FindSingleOfficeRelationship(relationships, relationshipSuffix, out bool multiple);
            if (match == null && !multiple) {
                return null;
            }
            if (multiple || match!.IsExternal) {
                throw new XlsxTabularFastPathNotSupportedException(
                    $"The workbook {relationshipName} relationship requires the Open XML SDK fallback path.");
            }

            string partName = ResolveTarget(workbookPartName, match.Target);
            if (!_parts.ContainsPart(partName)) {
                throw new InvalidDataException(
                    $"The workbook {relationshipName} part '{partName}' is missing.");
            }
            ValidatePartContentType(partName, expectedContentType, relationshipName);

            return partName;
        }

        private RichValueErrorLookup ReadRichValueErrors(string workbookPartName, Dictionary<string, PackageRelationship> relationships) {
            string? metadata = ResolveOptionalPart(workbookPartName, relationships, "/sheetMetadata", "cell metadata", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml");
            if (metadata == null) return RichValueErrorLookup.Empty;
            string? values = ResolveOptionalPart(workbookPartName, relationships, "/rdRichValue", "rich values", "application/vnd.ms-excel.rdrichvalue+xml");
            string? structures = ResolveOptionalPart(workbookPartName, relationships, "/rdRichValueStructure", "rich value structures", "application/vnd.ms-excel.rdrichvaluestructure+xml");
            if (values == null || structures == null) return RichValueErrorLookup.Empty;
            int maximumBytes = (int)Math.Min(_options.MaxInputBytes, Math.Min(_options.MaxMetadataPartBytes, RichValueErrorLookup.MaximumPartBytes));
            string limitName = _options.MaxInputBytes <= _options.MaxMetadataPartBytes
                && _options.MaxInputBytes <= RichValueErrorLookup.MaximumPartBytes
                ? nameof(ExcelReaderOptions.MaxInputBytes)
                : _options.MaxMetadataPartBytes <= RichValueErrorLookup.MaximumPartBytes
                    ? nameof(ExcelReaderOptions.MaxMetadataPartBytes)
                    : "MaxRichValuePartBytes";
            string Read(string name) {
                using Stream stream = _parts.OpenPart(name, maximumBytes, _options.CancellationToken, limitName: limitName);
                string xml = RichValueErrorLookup.ReadXml(stream);
                _options.CancellationToken.ThrowIfCancellationRequested();
                return xml;
            }
            return RichValueErrorLookup.FromRoots(
                new DocumentFormat.OpenXml.Spreadsheet.Metadata(Read(metadata)),
                new DocumentFormat.OpenXml.Office2019.Excel.RichData.RichValueData(Read(values)),
                new DocumentFormat.OpenXml.Office2019.Excel.RichData.RichValueStructures(Read(structures)));
        }

        private XDocument ReadXmlPart(string partName, int maximumBytes) {
            try {
                using Stream stream = _parts.OpenPart(partName, maximumBytes, _options.CancellationToken,
                    limitName: nameof(ExcelReaderOptions.MaxMetadataPartBytes));
                using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                    NameTable = new OpenXmlReadNameTable(),
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    CloseInput = false,
                    MaxCharactersInDocument = maximumBytes
                });
                XDocument result = XDocument.Load(reader, LoadOptions.None);
                _options.CancellationToken.ThrowIfCancellationRequested();
                return result;
            } catch (XmlException exception) {
                throw new XlsxTabularFastPathNotSupportedException(
                    $"Package part '{partName}' requires the Open XML SDK fallback path.",
                    exception);
            }
        }

        /// <summary>
        /// Parses a small package metadata part from its UTF-8 bytes, positioned on the root element. The rest of the
        /// part is read afterwards, so it must be well-formed as a whole. Parts the reader does not accept, and parts
        /// that declare another encoding for non-ASCII content, use the SDK path.
        /// </summary>
        private TResult ReadXmlPart<TResult>(string partName, int maximumBytes, Func<OfficeXmlPullReader, TResult> parse) {
            byte[]? rented = null;
            try {
                using Stream stream = _parts.OpenPart(partName, maximumBytes, _options.CancellationToken,
                    limitName: nameof(ExcelReaderOptions.MaxMetadataPartBytes));
                byte[] data;
                int length;
                if (stream is OpenXmlPooledPartStream pooled) {
                    data = pooled.BorrowBuffer(out length);
                } else {
                    // Streams of oversized or unusual entries need not know their length; OpenPart bounds what they return.
                    data = rented = ArrayPool<byte>.Shared.Rent(4096);
                    length = 0;
                    int read;
                    while ((read = stream.Read(data, length, data.Length - length)) > 0) {
                        length += read;
                        if (length < data.Length) continue;
                        byte[] larger = ArrayPool<byte>.Shared.Rent(checked(data.Length * 2));
                        Buffer.BlockCopy(data, 0, larger, 0, length);
                        ArrayPool<byte>.Shared.Return(data, clearArray: true);
                        data = rented = larger;
                    }
                }
                if (DeclaresOtherEncodingForNonAscii(data, length)) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        $"Package part '{partName}' requires the Open XML SDK fallback path.");
                }
                var reader = new OfficeXmlPullReader(data, length, _metadataNamespaces ??= new OfficeXmlNamespaces.Table(),
                    processMarkupCompatibility: false, recycle: _metadataReader);
                _metadataReader = reader;
                if (!reader.Read() || reader.NodeKind != OfficeXmlNodeKind.StartElement) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        $"Package part '{partName}' requires the Open XML SDK fallback path.");
                }
                TResult result = parse(reader);
                while (reader.Read()) {
                    _options.CancellationToken.ThrowIfCancellationRequested();
                }
                _options.CancellationToken.ThrowIfCancellationRequested();
                return result;
            } catch (XmlException exception) {
                throw new XlsxTabularFastPathNotSupportedException(
                    $"Package part '{partName}' requires the Open XML SDK fallback path.",
                    exception);
            } finally {
                if (rented != null) ArrayPool<byte>.Shared.Return(rented, clearArray: true);
            }
        }

        /// <summary>
        /// The pull reader reads UTF-8 (and UTF-16 by its byte order mark). An XML declaration naming another
        /// encoding only changes the meaning of non-ASCII bytes, so only such parts need the SDK path.
        /// </summary>
        private static bool DeclaresOtherEncodingForNonAscii(byte[] data, int length) {
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
            if (IsUtf8EncodingName(name)) return false;
            foreach (byte value in text) {
                if (value >= 0x80) return true;
            }
            return false;
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
            if (OpenXmlPartName.IsCanonicalPath(combined)) return combined;
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
                // Canonicalize both decoded ZIP entry names and encoded relationship targets
                // to the same idempotent OPC URI key.
                segments.Add(Uri.EscapeDataString(segment));
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

            if (OpenXmlPartName.IsCanonicalPath(partName, startIndex: 1)) return partName;
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
                    segments.Add(Uri.EscapeDataString(segment));
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

        private static bool ReadRelationshipTargetMode(string? targetMode) {
            if (targetMode == null || targetMode == "Internal") {
                return false;
            }
            if (targetMode == "External") {
                return true;
            }

            throw new XlsxTabularFastPathNotSupportedException(
                "A package relationship target mode requires the Open XML SDK fallback path.");
        }

        private static bool IsSupportedNonWorksheetRelationship(string relationshipType) =>
            IsOfficeRelationship(relationshipType, ChartSheetRelationshipSuffix)
            || IsOfficeRelationship(relationshipType, DialogSheetRelationshipSuffix)
            || IsOfficeRelationship(relationshipType, MacroSheetRelationshipSuffix)
            || IsOfficeRelationship(relationshipType, InternationalMacroSheetRelationshipSuffix)
            || string.Equals(relationshipType, OfficeMacroSheetRelationshipType, StringComparison.Ordinal)
            || string.Equals(relationshipType, OfficeInternationalMacroSheetRelationshipType, StringComparison.Ordinal);

        private static PackageRelationship? FindSingleOfficeRelationship(
            Dictionary<string, PackageRelationship> relationships, string suffix, out bool multiple) {
            PackageRelationship? match = null;
            multiple = false;
            foreach (KeyValuePair<string, PackageRelationship> relationship in relationships) {
                if (!IsOfficeRelationship(relationship.Value.Type, suffix)) continue;
                if (match != null) {
                    multiple = true;
                    return null;
                }
                match = relationship.Value;
            }
            return match;
        }

        private static bool IsOfficeRelationship(string? relationshipType, string suffix) =>
            relationshipType != null
            && relationshipType.EndsWith(suffix, StringComparison.Ordinal)
            && (relationshipType.Length == TransitionalOfficeRelationshipsNamespace.Length + suffix.Length
                    && relationshipType.StartsWith(TransitionalOfficeRelationshipsNamespace, StringComparison.Ordinal)
                || relationshipType.Length == StrictOfficeRelationshipsNamespace.Length + suffix.Length
                    && relationshipType.StartsWith(StrictOfficeRelationshipsNamespace, StringComparison.Ordinal));

        private static XlsxTabularSheet? ResolvePrefetchedSheet(
            IReadOnlyList<XlsxTabularSheet> sheets,
            ExcelReaderOptions options) {
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
                    try {
                        _sharedStrings.Dispose();
                    } finally {
                        _prefetchedParts?.Dispose();
                    }
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
