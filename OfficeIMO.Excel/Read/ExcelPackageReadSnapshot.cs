using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.IO.Packaging;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Read-only view of the package bytes an unchanged <see cref="ExcelDocument"/> was opened from.
    /// Reading part streams through the editable SDK package would invalidate its loaded roots, and
    /// loading those roots costs far more than the read, so readers use this view instead and the
    /// editable document stays untouched until something edits it.
    /// </summary>
    internal sealed class ExcelPackageReadSnapshot {
        private readonly MemoryStream _stream;
        private readonly Package _package;
        private readonly object _gate = new();
        private SharedStringCache? _parkedSharedStrings;
        private int _references = 1;

        private ExcelPackageReadSnapshot(
            byte[] bytes,
            MemoryStream stream,
            Package package,
            SpreadsheetDocument document,
            OpenXmlPackagePartBufferReader partBufferReader) {
            Bytes = bytes;
            _stream = stream;
            _package = package;
            Document = document;
            PartBufferReader = partBufferReader;
            Styles = new StylesCacheProvider(document);
        }

        internal byte[] Bytes { get; }

        internal SpreadsheetDocument Document { get; }

        internal OpenXmlPackagePartBufferReader PartBufferReader { get; }

        internal StylesCacheProvider Styles { get; }

        /// <summary>Returns null when the bytes need the editable document's more tolerant route.</summary>
        internal static ExcelPackageReadSnapshot? TryOpen(byte[] bytes) {
            MemoryStream? stream = null;
            Package? package = null;
            SpreadsheetDocument? document = null;
            OpenXmlPackagePartBufferReader? parts = null;
            try {
                parts = OpenXmlPackagePartBufferReader.TryOpen(bytes);
                if (parts == null) return null;
                stream = new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: false);
                package = Package.Open(stream, FileMode.Open, FileAccess.Read);
                document = SpreadsheetDocument.Open(package);
                return new ExcelPackageReadSnapshot(bytes, stream, package, document, parts);
            } catch (Exception exception) when (exception is IOException
                or InvalidDataException
                or InvalidOperationException
                or OpenXmlPackageException
                or NotSupportedException
                or ArgumentException) {
                parts?.Dispose();
                document?.Dispose();
                package?.Close();
                stream?.Dispose();
                return null;
            }
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
