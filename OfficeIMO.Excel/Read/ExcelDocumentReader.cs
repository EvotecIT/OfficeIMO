using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.OpenXml.Internal;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>Reads worksheets from a package or the live editing model.</summary>
    internal sealed partial class ExcelDocumentReader : IDisposable {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
        private const int Relationships = OfficeXmlNamespaces.Relationships;

        private readonly ExcelPartRef? _liveWorkbook;
        private readonly ExcelReaderPackage? _package;
        private readonly ExcelPackageReadSnapshot? _snapshot;
        private readonly IDisposable? _ownedResource;
        private readonly ExcelReaderOptions _opt;
        private readonly bool _qualifyWorksheets;
        private readonly bool _readOnlySnapshot;
        private readonly ExcelDateSystem _dateSystem;
        private readonly SharedStringCache _sst;
        private readonly StylesCacheProvider _styles;
        private readonly Lazy<RichValueErrorLookup> _richValueErrors;
        private SharedStringCache? _editedSharedStrings;
        private StylesCacheProvider? _editedStyles;
        private Lazy<RichValueErrorLookup>? _editedRichValueErrors;
        private int _disposed;

        private ExcelDocumentReader(ExcelPartRef? liveWorkbook, ExcelReaderPackage? package, ExcelReaderOptions options,
            ExcelPackageReadSnapshot? snapshot = null, IDisposable? ownedResource = null, bool tabular = false, bool readOnlySnapshot = false) {
            _liveWorkbook = liveWorkbook;
            _package = package;
            _snapshot = snapshot;
            _ownedResource = ownedResource;
            _opt = options;
            _qualifyWorksheets = tabular;
            _readOnlySnapshot = readOnlySnapshot;
            _opt.CancellationToken.ThrowIfCancellationRequested();
            _dateSystem = ReadWorkbookNode.Child(Spreadsheet, "workbookPr").BooleanAttribute("date1904") == true
                ? ExcelDateSystem.NineteenFour
                : ExcelDateSystem.NineteenHundred;
            _sst = snapshot?.RentSharedStrings(options) ?? package?.CreateSharedStrings(options, false)
                ?? SharedStringCache.Build(liveWorkbook.GetValueOrDefault(), options);
            _styles = snapshot?.Styles ?? package?.CreateStylesProvider() ?? new StylesCacheProvider(liveWorkbook.GetValueOrDefault());
            _richValueErrors = package == null
                ? new Lazy<RichValueErrorLookup>(() => RichValueErrorLookup.FromWorkbook(liveWorkbook.GetValueOrDefault(), RichValuePartLimit))
                : new Lazy<RichValueErrorLookup>(() => package.ReadRichValueErrors(_opt));
        }

        private int RichValuePartLimit => Math.Min(_opt.MaxMetadataPartBytes, RichValueErrorLookup.MaximumPartBytes);

        private OfficeXmlNode ReadWorkbookNode => _liveWorkbook is ExcelPartRef workbook ? workbook.Root : _package?.Workbook ?? default;

        public static ExcelDocumentReader Open(string path, ExcelReaderOptions? options = null) {
            ArgumentNullException.ThrowIfNull(path);
            if (!File.Exists(path)) throw new FileNotFoundException($"File '{path}' doesn't exist.", path);
            ExcelReaderOptions effective = options ?? new ExcelReaderOptions();
            SharedReadOnlyFileSnapshot? source = SharedReadOnlyFileSnapshot.Open(path);
            try {
                CheckInputLength(source.Length, effective);
                using (Stream directory = source.CreateView(bufferSize: 1)) {
                    ExcelDocumentLoadRouting.ValidateZipPackage(directory, effective.CancellationToken);
                }
                SharedReadOnlyFileSnapshot opened = source;
                try {
                    ExcelDocumentReader reader = OpenPackage(
                        () => opened.CreateView(), OpenXmlPackagePartBufferReader.TryOpen(opened.CreateView(bufferSize: 1), ignoreWhitespaceNames: true),
                        effective, opened);
                    source = null;
                    return reader;
                } catch (Exception ex) when (CanRetryWithNormalizedContentTypes(ex)) {
                    using Stream stream = opened.CreateView();
                    byte[] bytes = ReadInput(stream, effective);
                    return OpenFromBytes(bytes, effective, normalize: true, $"'{path}'");
                }
            } finally {
                source?.Dispose();
            }
        }

        public static ExcelDocumentReader Open(Stream stream, ExcelReaderOptions? options = null) {
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            ExcelReaderOptions effective = options ?? new ExcelReaderOptions();
            return OpenFromBytes(ReadInput(stream, effective), effective, normalize: false, "workbook stream");
        }

        public static ExcelDocumentReader Open(byte[] bytes, ExcelReaderOptions? options = null) {
            ArgumentNullException.ThrowIfNull(bytes);
            return OpenFromBytes(bytes, options ?? new ExcelReaderOptions(), normalize: false, "workbook bytes");
        }

        private static byte[] ReadInput(Stream stream, ExcelReaderOptions options) =>
            ExcelReadLimitFailure.ReadInput(() => OfficeStreamReader.ReadAllBytes(
                stream, options.CancellationToken, options.MaxInputBytes), options.MaxInputBytes);

        private static void CheckInputLength(long length, ExcelReaderOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (length > options.MaxInputBytes) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxInputBytes), options.MaxInputBytes,
                    $"Workbook input contains {length} bytes, exceeding the configured limit of {options.MaxInputBytes} bytes.");
            }
        }

        private static ExcelDocumentReader OpenFromBytes(byte[] bytes, ExcelReaderOptions options, bool normalize, string context) {
            CheckInputLength(bytes.LongLength, options);
            ExcelDocumentLoadRouting.ValidateZipPackage(bytes);
            try {
                if (normalize) {
                    using var stream = OpenNormalizedPackageStream(bytes, options.CancellationToken);
                    byte[] normalized = stream.GetBuffer();
                    int length = checked((int)stream.Length);
                    return OpenPackage(
                        () => new MemoryStream(normalized, 0, length, writable: false),
                        OpenXmlPackagePartBufferReader.TryOpen(new MemoryStream(normalized, 0, length, writable: false), ignoreWhitespaceNames: true), options);
                }

                return OpenPackage(() => new MemoryStream(bytes, writable: false),
                    OpenXmlPackagePartBufferReader.TryOpen(bytes, ignoreWhitespaceNames: true), options);
            } catch (Exception ex) when (CanRetryWithNormalizedContentTypes(ex)) {
                if (!normalize) return OpenFromBytes(bytes, options, normalize: true, context);
                throw new IOException($"Failed to open {context} after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'. See inner exception for details.", ex);
            }
        }

        private static MemoryStream OpenNormalizedPackageStream(byte[] bytes, CancellationToken cancellationToken) {
            var stream = new MemoryStream(bytes.Length + 4096);
            try {
                stream.Write(bytes, 0, bytes.Length);
                stream.Position = 0;
                ExcelPackageUtilities.NormalizeContentTypes(stream, leaveOpen: true);
                cancellationToken.ThrowIfCancellationRequested();
                stream.Position = 0;
                return stream;
            } catch {
                stream.Dispose();
                throw;
            }
        }

        private static bool CanRetryWithNormalizedContentTypes(Exception ex) =>
            !ExcelReadLimitFailure.Is(ex) && !ExcelPackagePartLengthFailure.Is(ex) && (ex is InvalidDataException or XmlException);

        private static ExcelDocumentReader OpenPackage(Func<Stream> open, OpenXmlPackagePartBufferReader? parts,
            ExcelReaderOptions options, IDisposable? ownedResource = null) {
            if (parts == null) throw ExcelPackageUtilities.UnreadablePackageContainer();
            ExcelReaderPackage? package = null;
            try {
                package = ExcelReaderPackage.Open(open, parts, options);
                return new ExcelDocumentReader(default, package, options, ownedResource: ownedResource);
            } catch {
                if (package == null) parts.Dispose();
                else package.Dispose();
                throw;
            }
        }

        public static ExcelDocumentReader Wrap(object document, ExcelReaderOptions? options = null) {
            ArgumentNullException.ThrowIfNull(document);
            return new ExcelDocumentReader(ExcelPartRef.FromSdk(document), null, options ?? new ExcelReaderOptions());
        }

        internal static ExcelDocumentReader WrapEdited(object document, ExcelPackageReadSnapshot baseline, ExcelReaderOptions options, bool tabular) {
            try {
                return new ExcelDocumentReader(ExcelPartRef.FromSdk(document), baseline.Package, options, baseline, tabular: tabular);
            } catch {
                baseline.Release();
                throw;
            }
        }

        internal static ExcelDocumentReader Wrap(ExcelPackageReadSnapshot snapshot, ExcelReaderOptions options, bool tabular = false,
            bool readOnlySnapshot = false) {
            try {
                return new ExcelDocumentReader(default, snapshot.Package, options, snapshot, tabular: tabular, readOnlySnapshot: readOnlySnapshot);
            } catch {
                snapshot.Release();
                throw;
            }
        }

        internal ExcelSheetReader GetSheet(string name, object worksheetPart) {
            ArgumentNullException.ThrowIfNull(worksheetPart);
            var part = ExcelPartRef.FromSdk(worksheetPart);
            return CreateSheetReader(name, part.Name, part);
        }

        private ExcelSheetReader CreateSheetReader(string sheetName, string partName, ExcelPartRef? livePart) {
            ExcelPartRef worksheet = livePart.GetValueOrDefault();
            ExcelPartRef workbook = _liveWorkbook.GetValueOrDefault();
            if (livePart == null || (_snapshot != null && !worksheet.IsRootLoaded && _package!.PartExists(partName))) {
                bool canStream = !_qualifyWorksheets || _snapshot!.QualifiesForTabularReaders(partName, _opt);
                if (canStream || livePart == null || (!workbook.WorkbookStyles.IsRootLoaded
                    && !workbook.FindWorkbookPart(ExcelReadWorkbookParts.SharedStrings).IsRootLoaded)) {
                    return new ExcelSheetReader(sheetName, partName, _sst, _styles, _opt, _dateSystem,
                        _package!.Parts, _richValueErrors, package: _package, canStreamWorksheetPart: canStream,
                        readOnlySnapshot: _readOnlySnapshot);
                }
            }

            SharedStringCache strings = _snapshot != null && workbook.FindWorkbookPart(ExcelReadWorkbookParts.SharedStrings).IsRootLoaded
                ? _editedSharedStrings ??= SharedStringCache.Build(workbook, _opt)
                : _sst;
            StylesCacheProvider styles = _snapshot != null && workbook.WorkbookStyles.IsRootLoaded
                ? _editedStyles ??= new StylesCacheProvider(workbook)
                : _styles;
            Lazy<RichValueErrorLookup> richValues = _snapshot == null ? _richValueErrors
                : _editedRichValueErrors ??= new Lazy<RichValueErrorLookup>(() => RichValueErrorLookup.FromWorkbook(workbook, RichValuePartLimit));
            return new ExcelSheetReader(sheetName, worksheet, strings, styles, _opt, _dateSystem,
                workbook.IsReadOnly, null, richValues);
        }

        public void Dispose() {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
            try {
                DisposeTableWorkbook();
                if (_snapshot != null) {
                    _editedSharedStrings?.Dispose();
                    _snapshot.ReturnSharedStrings(_sst);
                } else {
                    _sst.Dispose();
                }
            } finally {
                if (_snapshot != null) {
                    _snapshot.Release();
                } else {
                    try {
                        _package?.Dispose();
                    } finally {
                        _ownedResource?.Dispose();
                    }
                }
            }
        }
    }
}
