using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel {
    /// <summary>A shared read-only view of the package an unchanged workbook was opened from or last saved as.</summary>
    internal sealed partial class ExcelPackageReadSnapshot {
        private readonly IDisposable? _ownedFile;
        private readonly object _gate = new();
        private SharedStringCache? _parkedSharedStrings;
        private int _references = 1;

        private ExcelPackageReadSnapshot(object identity, ExcelReaderPackage package, IDisposable? ownedFile, StylesCache? styles) {
            Identity = identity;
            Package = package;
            _ownedFile = ownedFile;
            Styles = styles != null ? new StylesCacheProvider(styles) : package.CreateStylesProvider();
        }

        internal object Identity { get; }

        internal ExcelReaderPackage Package { get; }

        private OpenXmlPackagePartBufferReader PartBufferReader => Package.Parts;

        internal StylesCacheProvider Styles { get; }

        internal static ExcelPackageReadSnapshot? TryOpen(byte[] bytes, StylesCache? styles = null) =>
            TryOpen(bytes, OpenXmlPackagePartBufferReader.TryOpen(bytes, ignoreWhitespaceNames: true), null, styles);

        /// <summary>Copies a file-backed package before the editable workbook can rewrite its underlying file.</summary>
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
                ExcelPackageReadSnapshot? snapshot = TryOpen(identity,
                    OpenXmlPackagePartBufferReader.TryOpen(opened.CreateView(bufferSize: 1), ignoreWhitespaceNames: true), opened, styles: null);
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

        private static ExcelPackageReadSnapshot? TryOpen(object identity, OpenXmlPackagePartBufferReader? parts, IDisposable? ownedFile, StylesCache? styles) {
            if (parts == null) return null;
            try {
                return new ExcelPackageReadSnapshot(identity, ExcelReaderPackage.Open(parts), ownedFile, styles);
            } catch (Exception exception) when (exception is IOException or InvalidDataException or InvalidOperationException
                or System.Xml.XmlException or NotSupportedException or ArgumentException) {
                parts.Dispose();
                return null;
            } catch {
                parts.Dispose();
                throw;
            }
        }

        internal void AddReference() => Interlocked.Increment(ref _references);

        internal void Release() {
            if (Interlocked.Decrement(ref _references) != 0) return;
            lock (_gate) {
                _parkedSharedStrings?.Dispose();
                _parkedSharedStrings = null;
            }
            try {
                Package.Dispose();
            } finally {
                _ownedFile?.Dispose();
            }
        }

        /// <summary>Rents the parsed strings exclusively so concurrent readers have separate text caches.</summary>
        internal SharedStringCache RentSharedStrings(ExcelReaderOptions options) {
            lock (_gate) {
                SharedStringCache? parked = _parkedSharedStrings;
                _parkedSharedStrings = null;
                if (parked != null) {
                    if (parked.HasSameLimits(options)) return parked;
                    parked.Dispose();
                }
            }
            return Package.CreateSharedStrings(options, true);
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
