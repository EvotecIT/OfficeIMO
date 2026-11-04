using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using System.IO;
using System.IO.Packaging;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Read-only view of the package an unchanged <see cref="ExcelDocument"/> was opened from or last saved as.
    /// Reading part streams through the editable SDK package would invalidate its loaded roots, and
    /// loading those roots costs far more than the read, so readers use this view instead and the
    /// editable document stays untouched until something edits it.
    /// </summary>
    internal sealed partial class ExcelPackageReadSnapshot {
        private readonly Stream _stream;
        private readonly Package _package;
        private readonly IDisposable? _ownedFile;
        private readonly object _gate = new();
        private SharedStringCache? _parkedSharedStrings;
        private Dictionary<Uri, WorksheetPart>? _worksheetParts;
        private int _references = 1;

        private ExcelPackageReadSnapshot(
            object identity,
            Stream stream,
            Package package,
            SpreadsheetDocument document,
            OpenXmlPackagePartBufferReader partBufferReader,
            IDisposable? ownedFile,
            StylesCache? styles) {
            Identity = identity;
            _stream = stream;
            _package = package;
            _ownedFile = ownedFile;
            Document = document;
            PartBufferReader = partBufferReader;
            Styles = styles != null ? new StylesCacheProvider(styles) : new StylesCacheProvider(document);
        }

        /// <summary>The package version this view shows; a different value means it is out of date.</summary>
        internal object Identity { get; }

        internal SpreadsheetDocument Document { get; }

        internal OpenXmlPackagePartBufferReader PartBufferReader { get; }

        internal StylesCacheProvider Styles { get; }

        /// <summary>Returns null when the bytes need the editable document's more tolerant route.</summary>
        internal static ExcelPackageReadSnapshot? TryOpen(byte[] bytes, StylesCache? styles = null) =>
            TryOpen(
                bytes,
                () => OpenXmlPackagePartBufferReader.TryOpen(bytes),
                new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: false),
                ownedFile: null,
                styles);

        /// <summary>
        /// Copies a file-backed package to a private temporary file. That package file is rewritten in
        /// place by Save and allows no second handle, so readers cannot share it safely.
        /// </summary>
        internal static ExcelPackageReadSnapshot? TryOpenCopy(Stream package, object identity, CancellationToken cancellationToken) {
            FileStream? copy = null;
            SharedReadOnlyFileSnapshot? file = null;
            long position = package.Position;
            try {
                copy = OfficeTemporaryFile.Create("OfficeIMO.Excel-Read-", ".xlsx", FileOptions.RandomAccess, out _);
                package.Position = 0;
                package.CopyTo(copy);
                copy.Flush();
                cancellationToken.ThrowIfCancellationRequested();
                file = SharedReadOnlyFileSnapshot.Adopt(copy);
                copy = null;
                SharedReadOnlyFileSnapshot opened = file;
                ExcelPackageReadSnapshot? snapshot = TryOpen(
                    identity,
                    () => OpenXmlPackagePartBufferReader.TryOpen(opened.CreateView(bufferSize: 1)),
                    opened.CreateView(),
                    opened,
                    styles: null);
                if (snapshot != null) file = null;
                return snapshot;
            } catch (IOException) {
                return null;
            } catch (UnauthorizedAccessException) {
                return null;
            } finally {
                package.Position = position;
                copy?.Dispose();
                file?.Dispose();
            }
        }

        private static ExcelPackageReadSnapshot? TryOpen(
            object identity,
            Func<OpenXmlPackagePartBufferReader?> openParts,
            Stream stream,
            IDisposable? ownedFile,
            StylesCache? styles) {
            Package? package = null;
            SpreadsheetDocument? document = null;
            OpenXmlPackagePartBufferReader? parts = null;
            try {
                parts = openParts();
                if (parts != null) {
                    package = Package.Open(stream, FileMode.Open, FileAccess.Read);
                    document = SpreadsheetDocument.Open(package);
                    return new ExcelPackageReadSnapshot(identity, stream, package, document, parts, ownedFile, styles);
                }
            } catch (Exception exception) when (exception is IOException
                or InvalidDataException
                or InvalidOperationException
                or OpenXmlPackageException
                or NotSupportedException
                or ArgumentException) {
                parts?.Dispose();
                document?.Dispose();
                package?.Close();
            }

            stream.Dispose();
            return null;
        }

        internal void AddReference() => Interlocked.Increment(ref _references);

        internal void Release() {
            if (Interlocked.Decrement(ref _references) != 0) return;
            lock (_gate) {
                _parkedSharedStrings?.Dispose();
                _parkedSharedStrings = null;
            }
            Document.Dispose();
            _package.Close();
            _stream.Dispose();
            PartBufferReader.Dispose();
            _ownedFile?.Dispose();
        }

        /// <summary>Finds the part of this package an editable document knows by the same name.</summary>
        internal WorksheetPart? FindWorksheetPart(Uri uri) {
            lock (_gate) {
                _worksheetParts ??= Document.WorkbookPart?.WorksheetParts.ToDictionary(static part => part.Uri)
                    ?? new Dictionary<Uri, WorksheetPart>();
                return _worksheetParts.TryGetValue(uri, out WorksheetPart? part) ? part : null;
            }
        }

        /// <summary>
        /// Rents the parsed table exclusively, because its text caches are not shared between threads.
        /// Each sheet of a workbook then reuses one parse instead of repeating it.
        /// </summary>
        internal SharedStringCache RentSharedStrings(ExcelReaderOptions options) {
            lock (_gate) {
                SharedStringCache? parked = _parkedSharedStrings;
                _parkedSharedStrings = null;
                if (parked != null) {
                    if (parked.HasSameLimits(options)) return parked;
                    parked.Dispose();
                }
            }
            return SharedStringCache.BuildKeepingUnmarkedWhitespace(Document, options);
        }

        /// <summary>Keeps only a completed parse, so a canceled or failed load is never reused.</summary>
        internal void ReturnSharedStrings(SharedStringCache sharedStrings) {
            if (sharedStrings.TryDetachCancellationAfterLoad()) {
                SharedStringCache? replaced;
                lock (_gate) {
                    replaced = _parkedSharedStrings;
                    _parkedSharedStrings = sharedStrings;
                }
                replaced?.Dispose();
                return;
            }
            sharedStrings.Dispose();
        }
    }
}
