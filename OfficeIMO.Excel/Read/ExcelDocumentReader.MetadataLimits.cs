using System.Buffers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelDocumentReader {
        private static void ValidatePackageBootstrapMetadata(
            OpenXmlPackagePartBufferReader? partBufferReader,
            ExcelReadOptions options) {
            if (partBufferReader == null) {
                return;
            }

            ValidatePackageMetadataPartIfPresent(
                partBufferReader,
                "[Content_Types].xml",
                options);
            ValidatePackageMetadataPartIfPresent(
                partBufferReader,
                "_rels/.rels",
                options);
        }

        private void ValidateSdkMetadataLimits() {
            _opt.CancellationToken.ThrowIfCancellationRequested();
            if (!_owns) {
                // Opening a part stream on a mutable SDK document invalidates its loaded
                // root element. These limits protect package input, not the caller's
                // already-open authoring model.
                return;
            }

            WorkbookPart? workbookPart = _doc.WorkbookPart;
            if (workbookPart == null) {
                // A reader can wrap a newly-created workbook while its SDK root is still
                // being assembled. Package metadata limits apply when package metadata
                // exists; they must not make trusted in-memory authoring unreadable.
                return;
            }

            if (_partBufferReader != null
                && _partBufferReader.ContainsPart(workbookPart.Uri.OriginalString)) {
                ValidatePackageMetadataPartIfPresent(
                    _partBufferReader,
                    workbookPart.Uri.OriginalString,
                    _opt);
                ValidatePackageMetadataPartIfPresent(
                    _partBufferReader,
                    GetRelationshipsPartName(workbookPart.Uri),
                    _opt);
            } else {
                ValidateMetadataPartStream(workbookPart);
            }

            // The Workbook property materializes the SDK DOM. Keep that access after
            // raw package validation so an oversized compressed workbook part is
            // rejected before it can consume unbounded memory.
            Workbook? workbook = workbookPart.Workbook;
            if (workbook == null) {
                return;
            }

            int sheetDefinitions = 0;
            IEnumerable<Sheet> sheets =
                workbook.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>();
            foreach (Sheet _ in sheets) {
                _opt.CancellationToken.ThrowIfCancellationRequested();
                sheetDefinitions++;
                if (sheetDefinitions > _opt.MaxWorksheets) {
                    throw ExcelReadLimitFailure.Create(
                        $"The OpenXML workbook contains more than the configured {_opt.MaxWorksheets} worksheet definitions.");
                }
            }
        }

        private static void ValidatePackageMetadataPartIfPresent(
            OpenXmlPackagePartBufferReader partBufferReader,
            string partName,
            ExcelReadOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!partBufferReader.ContainsPart(partName)) {
                return;
            }

            using Stream stream = partBufferReader.OpenPart(
                partName,
                options.MaxMetadataPartBytes);
            DrainMetadataPartStream(stream, partName, options);
        }

        private void ValidateMetadataPartStream(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            DrainMetadataPartStream(stream, part.Uri.OriginalString, _opt);
        }

        internal static void DrainMetadataPartStream(
            Stream stream,
            string partName,
            ExcelReadOptions options) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (options == null) throw new ArgumentNullException(nameof(options));

            options.CancellationToken.ThrowIfCancellationRequested();
            if (stream.CanSeek && stream.Length > options.MaxMetadataPartBytes) {
                throw ExcelReadLimitFailure.Create(
                    $"Package part '{partName}' contains {stream.Length} bytes, exceeding the supported limit of {options.MaxMetadataPartBytes} bytes.");
            }

            int bufferLength = Math.Min(8192, options.MaxMetadataPartBytes);
            byte[] buffer = ArrayPool<byte>.Shared.Rent(bufferLength);
            try {
                long totalBytes = 0;
                int read;
                while ((read = stream.Read(buffer, 0, bufferLength)) > 0) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    totalBytes += read;
                    if (totalBytes > options.MaxMetadataPartBytes) {
                        throw ExcelReadLimitFailure.Create(
                            $"Package part '{partName}' exceeds the supported limit of {options.MaxMetadataPartBytes} bytes.");
                    }
                }
            } finally {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        private static string GetRelationshipsPartName(Uri partUri) {
            string partName = partUri.OriginalString.TrimStart('/');
            int separator = partName.LastIndexOf('/');
            if (separator < 0) {
                return $"_rels/{partName}.rels";
            }

            return $"{partName.Substring(0, separator + 1)}_rels/{partName.Substring(separator + 1)}.rels";
        }

        private static OpenXmlPackagePartBufferReader? TryOpenPartBufferReader(
            MemoryStream packageStream) {
            if (!packageStream.TryGetBuffer(out ArraySegment<byte> buffer)
                || buffer.Array == null) {
                return null;
            }

            var view = new MemoryStream(
                buffer.Array,
                buffer.Offset,
                checked((int)packageStream.Length),
                writable: false,
                publiclyVisible: false);
            return OpenXmlPackagePartBufferReader.TryOpen(view);
        }
    }
}
