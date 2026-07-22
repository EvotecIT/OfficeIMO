using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>Inspects the attached VBA project without executing macro code.</summary>
        /// <param name="includeSha256">Whether to calculate the project SHA-256 digest.</param>
        /// <returns>Project metadata, or null when the document has no VBA project.</returns>
        public OfficeVbaProjectInfo? InspectVbaProject(bool includeSha256 = false) {
            return InspectVbaProject(includeSha256, OfficeVbaProjectInfo.DefaultMaximumProjectBytes);
        }

        /// <summary>Inspects the attached VBA project while enforcing a decoded byte limit.</summary>
        /// <param name="includeSha256">Whether to calculate the project SHA-256 digest.</param>
        /// <param name="maxBytes">Maximum decoded VBA project bytes accepted.</param>
        /// <returns>Project metadata, or null when the document has no VBA project.</returns>
        public OfficeVbaProjectInfo? InspectVbaProject(bool includeSha256, long maxBytes) {
            if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            VbaProjectPart? part = _wordprocessingDocument.MainDocumentPart?.VbaProjectPart;
            return part == null ? null : OfficeOpenXmlPackagePayload.CreateVbaProjectInfo(
                part, includeSha256, maxBytes);
        }

        /// <summary>Extracts a VBA project while enforcing a maximum byte count.</summary>
        public byte[] ExtractMacros(long maxBytes) {
            if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            VbaProjectPart? part = _wordprocessingDocument.MainDocumentPart?.VbaProjectPart;
            return part == null ? Array.Empty<byte>() : OfficeOpenXmlPackagePayload.ReadBytes(part, maxBytes);
        }

        /// <summary>Saves a VBA project while enforcing a maximum byte count.</summary>
        public void SaveMacros(string filePath, long maxBytes) {
            if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            VbaProjectPart? part = _wordprocessingDocument.MainDocumentPart?.VbaProjectPart;
            if (part != null) {
                OfficeOpenXmlPackagePayload.SaveBytes(part, filePath, maxBytes);
            }
        }

        /// <summary>Adds or replaces a VBA project from a readable stream.</summary>
        /// <param name="stream">Stream containing <c>vbaProject.bin</c>.</param>
        /// <param name="maxBytes">Optional maximum number of bytes accepted from the stream.</param>
        public void AddMacro(Stream stream, long? maxBytes = null) {
            AddMacro(OfficeStreamReader.ReadAllBytes(stream, maxBytes));
        }

        /// <summary>Gets embedded package, OLE, and ActiveX payload metadata.</summary>
        /// <param name="includeSha256">Whether to calculate a digest for every payload.</param>
        public IReadOnlyList<OfficeEmbeddedPayloadInfo> GetEmbeddedPayloads(bool includeSha256 = false) {
            return OfficeOpenXmlPackagePayload.FindEmbeddedPayloads(_wordprocessingDocument)
                .Select(handle => OfficeOpenXmlPackagePayload.CreateInfo(handle, includeSha256))
                .ToArray();
        }

        /// <summary>Extracts an embedded payload by its package-local id.</summary>
        /// <param name="id">Id returned by <see cref="GetEmbeddedPayloads(bool)"/>.</param>
        /// <param name="maxBytes">Optional extraction limit.</param>
        public byte[] ExtractEmbeddedPayload(string id, long? maxBytes = null) {
            OfficeOpenXmlPayloadHandle handle = OfficeOpenXmlPackagePayload.FindEmbeddedPayload(_wordprocessingDocument, id);
            return OfficeOpenXmlPackagePayload.ReadBytes(handle.Part, maxBytes);
        }

        /// <summary>Saves an embedded payload by its package-local id.</summary>
        public void SaveEmbeddedPayload(string id, string filePath, long? maxBytes = null) {
            OfficeOpenXmlPayloadHandle handle = OfficeOpenXmlPackagePayload.FindEmbeddedPayload(_wordprocessingDocument, id);
            OfficeOpenXmlPackagePayload.SaveBytes(handle.Part, filePath, maxBytes);
        }

        /// <summary>Replaces an embedded payload while preserving its owner relationship.</summary>
        public void ReplaceEmbeddedPayload(string id, byte[] data) {
            OfficeOpenXmlPayloadHandle handle = OfficeOpenXmlPackagePayload.FindEmbeddedPayload(_wordprocessingDocument, id);
            OfficeOpenXmlPackagePayload.ReplaceBytes(handle.Part, data);
        }

        /// <summary>Replaces an embedded payload from a readable stream.</summary>
        public void ReplaceEmbeddedPayload(string id, Stream stream, long? maxBytes = null) {
            ReplaceEmbeddedPayload(id, OfficeStreamReader.ReadAllBytes(stream, maxBytes));
        }

        /// <summary>Removes an embedded payload and known Word OLE/control markup that references it.</summary>
        public bool RemoveEmbeddedPayload(string id) {
            OfficeOpenXmlPayloadHandle? handle = OfficeOpenXmlPackagePayload.FindEmbeddedPayloads(_wordprocessingDocument)
                .FirstOrDefault(candidate => string.Equals(candidate.Id, id, StringComparison.Ordinal));
            if (handle == null) {
                return false;
            }

            OfficeOpenXmlPackagePayload.RemovePart(handle);
            return true;
        }
    }
}
