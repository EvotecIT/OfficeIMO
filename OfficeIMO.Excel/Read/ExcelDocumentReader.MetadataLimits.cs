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
            WorkbookPart workbookPart = WorkbookPartRoot;

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

            int sheetDefinitions = 0;
            IEnumerable<Sheet> sheets =
                WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>();
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
            options.CancellationToken.ThrowIfCancellationRequested();
        }

        private void ValidateMetadataPartStream(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            if (stream.CanSeek && stream.Length > _opt.MaxMetadataPartBytes) {
                throw ExcelReadLimitFailure.Create(
                    $"Package part '{part.Uri}' contains {stream.Length} bytes, exceeding the supported limit of {_opt.MaxMetadataPartBytes} bytes.");
            }

            int bufferLength = Math.Min(8192, _opt.MaxMetadataPartBytes);
            byte[] buffer = new byte[bufferLength];
            long totalBytes = 0;
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
                _opt.CancellationToken.ThrowIfCancellationRequested();
                totalBytes += read;
                if (totalBytes > _opt.MaxMetadataPartBytes) {
                    throw ExcelReadLimitFailure.Create(
                        $"Package part '{part.Uri}' exceeds the supported limit of {_opt.MaxMetadataPartBytes} bytes.");
                }
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
