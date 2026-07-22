using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Indicates whether the workbook contains a VBA project.</summary>
        public bool HasMacros => WorkbookPartRoot?.VbaProjectPart != null;

        /// <summary>Inspects the attached VBA project without executing macro code.</summary>
        public OfficeVbaProjectInfo? InspectVbaProject(bool includeSha256 = false) {
            return InspectVbaProject(includeSha256, OfficeVbaProjectInfo.DefaultMaximumProjectBytes);
        }

        /// <summary>Inspects the attached VBA project while enforcing a decoded byte limit.</summary>
        public OfficeVbaProjectInfo? InspectVbaProject(bool includeSha256, long maxBytes) {
            if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            return Locking.ExecuteRead(EnsureLock(), () => {
                VbaProjectPart? part = WorkbookPartRoot?.VbaProjectPart;
                return part == null ? null : OfficeOpenXmlPackagePayload.CreateVbaProjectInfo(
                    part, includeSha256, maxBytes);
            });
        }

        /// <summary>Adds or replaces a <c>vbaProject.bin</c> payload.</summary>
        public void AddMacro(byte[] data) {
            if (data == null || data.Length == 0) throw new ArgumentException("VBA project data cannot be empty.", nameof(data));
            Locking.ExecuteWrite(EnsureLock(), () => {
                WorkbookPart workbookPart = WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is missing.");
                if (workbookPart.VbaProjectPart != null) {
                    workbookPart.DeletePart(workbookPart.VbaProjectPart);
                }

                VbaProjectPart part = workbookPart.AddNewPart<VbaProjectPart>();
                OfficeOpenXmlPackagePayload.ReplaceBytes(part, data);
                EnsureMacroEnabledDocumentType();
                MarkPackageDirty();
            });
        }

        /// <summary>Adds or replaces a VBA project from a file.</summary>
        public void AddMacro(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("VBA project path cannot be empty.", nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            AddMacro(File.ReadAllBytes(filePath));
        }

        /// <summary>Adds or replaces a VBA project from a readable stream.</summary>
        public void AddMacro(Stream stream, long? maxBytes = null) {
            AddMacro(OfficeStreamReader.ReadAllBytes(stream, maxBytes));
        }

        /// <summary>Extracts the VBA project, or an empty array when no project exists.</summary>
        public byte[] ExtractMacros(long? maxBytes = null) {
            return Locking.ExecuteRead(EnsureLock(), () => {
                VbaProjectPart? part = WorkbookPartRoot?.VbaProjectPart;
                return part == null ? Array.Empty<byte>() : OfficeOpenXmlPackagePayload.ReadBytes(part, maxBytes);
            });
        }

        /// <summary>Saves the VBA project when one is present.</summary>
        public void SaveMacros(string filePath, long? maxBytes = null) {
            Locking.ExecuteRead(EnsureLock(), () => {
                VbaProjectPart? part = WorkbookPartRoot?.VbaProjectPart;
                if (part != null) {
                    OfficeOpenXmlPackagePayload.SaveBytes(part, filePath, maxBytes);
                }
            });
        }

        /// <summary>Removes the complete VBA project.</summary>
        public void RemoveMacros() {
            Locking.ExecuteWrite(EnsureLock(), () => {
                WorkbookPart? workbookPart = WorkbookPartRoot;
                if (workbookPart?.VbaProjectPart == null) {
                    return;
                }

                workbookPart.DeletePart(workbookPart.VbaProjectPart);
                MarkPackageDirty();
            });
        }

        /// <summary>Gets embedded package, OLE, and ActiveX payload metadata.</summary>
        public IReadOnlyList<OfficeEmbeddedPayloadInfo> GetEmbeddedPayloads(bool includeSha256 = false) {
            return Locking.ExecuteRead(EnsureLock(), () =>
                OfficeOpenXmlPackagePayload.FindEmbeddedPayloads(_spreadSheetDocument)
                    .Select(handle => OfficeOpenXmlPackagePayload.CreateInfo(handle, includeSha256))
                    .ToArray());
        }

        /// <summary>Extracts an embedded payload by its package-local id.</summary>
        public byte[] ExtractEmbeddedPayload(string id, long? maxBytes = null) {
            return Locking.ExecuteRead(EnsureLock(), () => {
                OfficeOpenXmlPayloadHandle handle = OfficeOpenXmlPackagePayload.FindEmbeddedPayload(_spreadSheetDocument, id);
                return OfficeOpenXmlPackagePayload.ReadBytes(handle.Part, maxBytes);
            });
        }

        /// <summary>Saves an embedded payload by its package-local id.</summary>
        public void SaveEmbeddedPayload(string id, string filePath, long? maxBytes = null) {
            Locking.ExecuteRead(EnsureLock(), () => {
                OfficeOpenXmlPayloadHandle handle = OfficeOpenXmlPackagePayload.FindEmbeddedPayload(_spreadSheetDocument, id);
                OfficeOpenXmlPackagePayload.SaveBytes(handle.Part, filePath, maxBytes);
            });
        }

        /// <summary>Replaces an embedded payload while preserving its owner relationship.</summary>
        public void ReplaceEmbeddedPayload(string id, byte[] data) {
            Locking.ExecuteWrite(EnsureLock(), () => {
                OfficeOpenXmlPayloadHandle handle = OfficeOpenXmlPackagePayload.FindEmbeddedPayload(_spreadSheetDocument, id);
                OfficeOpenXmlPackagePayload.ReplaceBytes(handle.Part, data);
                MarkPackageDirty();
            });
        }

        /// <summary>Replaces an embedded payload from a readable stream.</summary>
        public void ReplaceEmbeddedPayload(string id, Stream stream, long? maxBytes = null) {
            ReplaceEmbeddedPayload(id, OfficeStreamReader.ReadAllBytes(stream, maxBytes));
        }

        /// <summary>Removes an embedded payload and known worksheet OLE/control markup that references it.</summary>
        public bool RemoveEmbeddedPayload(string id) {
            return Locking.ExecuteWrite(EnsureLock(), () => {
                OfficeOpenXmlPayloadHandle? handle = OfficeOpenXmlPackagePayload.FindEmbeddedPayloads(_spreadSheetDocument)
                    .FirstOrDefault(candidate => string.Equals(candidate.Id, id, StringComparison.Ordinal));
                if (handle == null) {
                    return false;
                }

                OfficeOpenXmlPackagePayload.RemovePart(handle);
                MarkPackageDirty();
                return true;
            });
        }

        private void EnsureMacroEnabledDocumentType() {
            SpreadsheetDocumentType type = _spreadSheetDocument.DocumentType;
            SpreadsheetDocumentType target = type == SpreadsheetDocumentType.Template
                || type == SpreadsheetDocumentType.MacroEnabledTemplate
                ? SpreadsheetDocumentType.MacroEnabledTemplate
                : SpreadsheetDocumentType.MacroEnabledWorkbook;
            if (type != target) {
                _spreadSheetDocument.ChangeDocumentType(target);
            }
        }
    }
}
