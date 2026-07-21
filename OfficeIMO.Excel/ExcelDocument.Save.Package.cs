using OfficeIMO.Drawing.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        private void PrepareWorkbookForSave(ExcelSaveOptions? options, bool skipDirectFastSaveSheetPreparation = false) {
            ApplySignatureMutationPolicy(options);
            Stopwatch? stageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            if (skipDirectFastSaveSheetPreparation) {
                MaterializePendingDirectCellValueSheetIfNeeded();
            } else {
                MaterializeDeferredDataSetImport();
            }
            ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.MaterializeDeferredDataSet");
            using var preserveFastSaveState = _materializedDirectDataSetFastSaveModel != null
                ? PreserveDirectDataSetFastSaveStateDuringDirtyMarks()
                : null;

            // Ensure all worksheets have up-to-date dimensions and proper element ordering before saving
            ApplyCalculationPolicyBeforeSave(options);
            ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.ApplyCalculationPolicy");

            var sheets = Sheets;
            ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.GetSheets");
            foreach (var sheet in sheets) {
                if (!sheet.RequiresSavePreparation) {
                    continue;
                }

                if (skipDirectFastSaveSheetPreparation && IsMaterializedDirectDataSetFastSaveSheet(sheet)) {
                    ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.SkipDirectFastSaveSheet");
                    continue;
                }

                sheet.UpdateSheetDimension();
                ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.UpdateSheetDimension");
                sheet.EnsureWorksheetElementOrder();
                ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.EnsureWorksheetElementOrder");
                sheet.Commit();
                ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.CommitSheet");
            }

            // Run the heavier repair sweep only when workbook-level operations requested it.
            if (_requiresSavePreflight || options?.SafePreflight == true) {
                if (!skipDirectFastSaveSheetPreparation || options?.SafePreflight == true) {
                    MaterializePendingDirectCellValueSheetIfNeeded();
                    PreflightWorkbook(sheets);
                    _requiresSavePreflight = false;
                } else {
                    ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.SkipAutomaticPreflight");
                }
            }
            ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.Preflight");
            if (options?.SafePreflight == true) {
                // Already performed above; branch kept for semantic clarity
            }

            if (options?.SafeRepairDefinedNames == true) {
                RepairDefinedNames(save: true);
            }
            ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.RepairDefinedNames");

            if (_sharedStringTableDirty) {
                _sharedStringTablePart?.SharedStringTable?.Save();
                _sharedStringTableDirty = false;
            }
            ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.SaveSharedStrings");

            SaveCustomDocumentProperties();
            ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.SaveCustomDocumentProperties");

            WorkbookRoot.Save();
            ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.SaveWorkbookRoot");
            _spreadSheetDocument.PackageProperties.Modified = DateTime.UtcNow;
            ReportSaveTiming(stageWatch, "Save.PrepareWorkbook.UpdatePackageProperties");
        }

        private bool IsMaterializedDirectDataSetFastSaveSheet(ExcelSheet sheet) {
            var model = _materializedDirectDataSetFastSaveModel;
            if (model == null) {
                return false;
            }

            for (int i = 0; i < model.Sheets.Count; i++) {
                if (string.Equals(model.Sheets[i].SheetName, sheet.Name, StringComparison.Ordinal)) {
                    return true;
                }
            }

            return false;
        }

        private void ReportSaveTiming(Stopwatch? stopwatch, string operation) {
            if (stopwatch == null) {
                return;
            }

            Execution.ReportTiming(operation, stopwatch.Elapsed);
            stopwatch.Restart();
        }

        private SavePayload PreparePackageForSave(ExcelSaveOptions? options, bool closeDocument = true) {
            PrepareWorkbookForSave(options);

            PackagePropertiesSnapshot propertiesSnapshot = PackagePropertiesSnapshot.Capture(_spreadSheetDocument);

            using var snapshot = new MemoryStream();
            using (_spreadSheetDocument.Clone(snapshot)) { }
            snapshot.Position = 0;

            var packageBytes = snapshot.ToArray();

            if (closeDocument) {
                try { _spreadSheetDocument.Dispose(); } catch { }
            }

            return new SavePayload(packageBytes, propertiesSnapshot, closeDocument, normalizeContentTypes: !_packageContentTypesKnownNormalized, applyPackageProperties: _packagePropertiesDirty);
        }

        private static void PrepareDestinationStreamForWrite(Stream destination) {
            if (!destination.CanSeek) {
                return;
            }

            destination.Seek(0, SeekOrigin.Begin);
            destination.SetLength(0);
        }

        private bool TryWriteUnchangedPackageToStream(Stream destination, ExcelSaveOptions? options) {
            if (!CanUseUnchangedPackageFastPath(options) || _unchangedPackageBytes == null) {
                return false;
            }

            PrepareDestinationStreamForWrite(destination);
            destination.Write(_unchangedPackageBytes, 0, _unchangedPackageBytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            return true;
        }

        private async Task<bool> TryWriteUnchangedPackageToStreamAsync(Stream destination, ExcelSaveOptions? options, CancellationToken cancellationToken) {
            if (!CanUseUnchangedPackageFastPath(options) || _unchangedPackageBytes == null) {
                return false;
            }

            PrepareDestinationStreamForWrite(destination);
            await destination.WriteAsync(_unchangedPackageBytes, 0, _unchangedPackageBytes.Length, cancellationToken).ConfigureAwait(false);
            try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
            return true;
        }

        private bool CanUseUnchangedPackageFastPath(ExcelSaveOptions? options) {
            return !_packageDirty
                && _packageContentTypesKnownNormalized
                && _unchangedPackageBytes != null
                && !HasCalculationSaveWork(options)
                && options?.SafePreflight != true
                && options?.SafeRepairDefinedNames != true
                && options?.ValidateOpenXml != true;
        }

        private void MarkPackageClean(byte[]? packageBytes, bool simplePackageContentKnown = false) {
            _packageDirty = false;
            _packagePropertiesDirty = false;
            _unchangedPackageBytes = packageBytes;
            _packageContentTypesKnownNormalized = true;
            _simplePackageContentKnown = simplePackageContentKnown;
            _requiresSavePreflight = false;
        }

        private static string CreateTemporarySavePath(string targetPath) {
            return OfficeFileCommit.CreateTemporaryPath(targetPath);
        }

        private static void ReplaceTargetFile(string temporaryPath, string targetPath) {
            OfficeFileCommit.CommitTemporaryFile(temporaryPath, targetPath);
        }

        private static void DeleteFileIfExists(string path) {
            OfficeFileCommit.DeleteIfExists(path);
        }

        private bool TrySaveWithSimplePackageToFile(string targetPath, ExcelSaveOptions? options, out string? skipReason, CancellationToken ct = default, bool alreadyPrepared = false) {
            skipReason = null;
            var temporaryPath = CreateTemporarySavePath(targetPath);
            byte[]? packageBytes = null;

            try {
                if (!alreadyPrepared) {
                    PrepareWorkbookForSave(options);
                }

                using (var fs = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)) {
                    if (!TryWriteSimpleWorkbookPackage(fs, options, updateDocumentState: false, out skipReason, ct)) {
                        return false;
                    }
                }

                ct.ThrowIfCancellationRequested();
                packageBytes = File.ReadAllBytes(temporaryPath);

                try { _spreadSheetDocument.Dispose(); } catch { }
                ReplaceTargetFile(temporaryPath, targetPath);
                temporaryPath = string.Empty;
                ReloadFromBytes(packageBytes, simplePackageContentKnown: true);

                FilePath = targetPath;
                LastSaveDiagnostics = ExcelSaveDiagnostics.SimplePackage();
                return true;
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                skipReason = "Simple package writer failed: " + ex.Message;
                if (packageBytes != null) {
                    try { ReloadFromBytes(packageBytes, simplePackageContentKnown: true); } catch { }
                }

                return false;
            } finally {
                DeleteFileIfExists(temporaryPath);
            }
        }

        private bool TrySaveWithExtendedPackageToFile(string targetPath, ExcelSaveOptions? options, out string? skipReason, CancellationToken ct = default) {
            skipReason = null;
            var temporaryPath = CreateTemporarySavePath(targetPath);
            byte[]? packageBytes = null;

            try {
                using (var fs = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)) {
                    if (!TryWriteExtendedWorkbookPackage(fs, options, updateDocumentState: false, out skipReason, ct)) {
                        return false;
                    }
                }

                ct.ThrowIfCancellationRequested();
                packageBytes = File.ReadAllBytes(temporaryPath);

                try { _spreadSheetDocument.Dispose(); } catch { }
                ReplaceTargetFile(temporaryPath, targetPath);
                temporaryPath = string.Empty;
                ReloadFromBytes(packageBytes, simplePackageContentKnown: true);

                FilePath = targetPath;
                LastSaveDiagnostics = ExcelSaveDiagnostics.ExtendedPackage();
                return true;
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                skipReason = "Extended package writer failed: " + ex.Message;
                if (packageBytes != null) {
                    try { ReloadFromBytes(packageBytes, simplePackageContentKnown: true); } catch { }
                }

                return false;
            } finally {
                DeleteFileIfExists(temporaryPath);
            }
        }

        private static void CommitPreparedPackageToFile(string targetPath, byte[] finalizedBytes) {
            OfficeFileCommit.WriteAllBytes(targetPath, finalizedBytes);
        }

        private static Task CommitPreparedPackageToFileAsync(string targetPath, byte[] finalizedBytes, CancellationToken cancellationToken) {
            return OfficeFileCommit.WriteAllBytesAsync(targetPath, finalizedBytes, cancellationToken: cancellationToken);
        }

        private void TryRestoreDocumentState(SavePayload payload) {
            if (!payload.DocumentClosed) {
                return;
            }

            try {
                ReloadFromBytes(payload.PackageBytes);
            } catch {
                // Best-effort recovery only; preserve the original save exception.
            }
        }

        private void ReloadFromBytes(byte[] packageBytes, bool simplePackageContentKnown = false, Stream? reusablePackageStream = null) {
            var previousDocument = _spreadSheetDocument;
            var previousPackageStream = _packageStream;
            bool keepPackageStream = _copyPackageToSourceOnDispose || _copyPackageToFilePathOnDispose;
            bool previousDocumentDisposed = false;

            Stream mem;
            if (reusablePackageStream != null) {
                if (!ReferenceEquals(reusablePackageStream, previousPackageStream) || !keepPackageStream) {
                    throw new InvalidOperationException("Only the stream backing the current package can be reused.");
                }

                // Close the old package before replacing its bytes. This is required by the
                // .NET Framework packaging backend, which retains offsets into the open ZIP.
                previousDocument.Dispose();
                previousDocumentDisposed = true;
                PrepareDestinationStreamForWrite(reusablePackageStream);
                reusablePackageStream.Write(packageBytes, 0, packageBytes.Length);
                reusablePackageStream.Flush();
                reusablePackageStream.Position = 0;
                mem = reusablePackageStream;
            } else {
                mem = keepPackageStream
                    ? new NonDisposingMemoryStream(packageBytes.Length + 8192)
                    : new MemoryStream(packageBytes.Length + 8192);
                mem.Write(packageBytes, 0, packageBytes.Length);
                mem.Position = 0;
            }

            var reopenSettings = new OpenSettings { AutoSave = false };
            _spreadSheetDocument = SpreadsheetDocument.Open(mem, true, reopenSettings);
            _workBookPart = WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            _sharedStringTablePart = null;
            _sharedStringCache.Clear();
            _sharedStringTableCount = -1;
            _sharedStringTableDirty = false;
            _cachedSheets = null;
            _sheetCacheDirty = true;
            _packageStream = keepPackageStream ? mem : null;
            ReinitializePackageBoundHelpers();
            MarkPackageClean(packageBytes, simplePackageContentKnown);

            if (previousPackageStream != null && !ReferenceEquals(previousPackageStream, mem)) {
                DisposeStream(previousPackageStream);
            }

            if (!previousDocumentDisposed && previousDocument != null && !ReferenceEquals(previousDocument, _spreadSheetDocument)) {
                try { previousDocument.Dispose(); } catch { }
            }
        }

        private void ReinitializePackageBoundHelpers() {
            BuiltinDocumentProperties = new BuiltinDocumentProperties(this);
            ApplicationProperties = new ApplicationProperties(this);
            CustomDocumentProperties.SetChangeHandler(MarkCustomDocumentPropertiesChanged);
            LoadCustomDocumentProperties();
            ExcelChartAxisIdGenerator.Initialize(_spreadSheetDocument);
        }

        private static byte[] NormalizePackageBytes(byte[] packageBytes) {
            using (var probe = new MemoryStream(packageBytes, writable: false)) {
                try {
                    if (!ExcelPackageUtilities.NeedsContentTypeNormalization(probe)) {
                        return packageBytes;
                    }
                } catch {
                }
            }

            var working = new MemoryStream(packageBytes.Length + StreamBufferSize);
            working.Write(packageBytes, 0, packageBytes.Length);
            working.Position = 0;

            try {
                ExcelPackageUtilities.NormalizeContentTypes(working, leaveOpen: true);
            } catch {
            }

            if (working.CanSeek) {
                working.Position = 0;
            }

            return working.ToArray();
        }

        private static byte[] FinalizePackageBytes(SavePayload payload) {
            var withProperties = payload.ApplyPackageProperties
                ? payload.Properties.ApplyTo(payload.PackageBytes)
                : payload.PackageBytes;
            return payload.NormalizeContentTypes
                ? NormalizePackageBytes(withProperties)
                : withProperties;
        }

        private static void ThrowIfOpenXmlValidationFails(byte[] finalizedBytes, ExcelSaveOptions? options) {
            if (options?.ValidateOpenXml != true) {
                return;
            }

            var errors = ValidateOpenXml(finalizedBytes);
            if (errors.Count > 0) {
                throw new InvalidOperationException("OpenXML validation failed:\n" + string.Join("\n", errors));
            }
        }
    }
}
