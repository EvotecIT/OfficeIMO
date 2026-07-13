using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Drawing;
using OfficeIMO.Shared;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        private bool _disposed;

        /// <summary>
        /// Releases resources used by the document.
        /// </summary>
        public void Dispose() {
            if (_disposed) {
                return;
            }

            Exception? persistenceFailure = null;

            try {
                if (this._spreadSheetDocument != null) {
                    try {
                        if (_persistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                            bool shouldSave = this._spreadSheetDocument.FileOpenAccess != FileAccess.Read &&
                                              (IsPackageDirty || _copyPackageToSourceOnDispose || _copyPackageToFilePathOnDispose);
                            _copyPackageToSourceOnDispose = false;
                            _copyPackageToFilePathOnDispose = false;
                            if (shouldSave) {
                                Save();
                            }
                        }

                        this._spreadSheetDocument.Dispose();
                    } catch (Exception ex) {
                        persistenceFailure = ex;
                    } finally {
                        this._spreadSheetDocument = null!;
                    }
                }

                try {
                    PersistPackageToSourceIfNeeded();
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }
            } finally {
                if (_ownedOpenStream != null) {
                    try {
                        _ownedOpenStream.Dispose();
                    } catch (Exception ex) {
                        persistenceFailure ??= ex;
                    }
                    _ownedOpenStream = null;
                }

                _lock?.Dispose();
                _disposed = true;
                GC.SuppressFinalize(this);
            }

            if (persistenceFailure != null) {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(persistenceFailure).Throw();
            }
        }

        /// <summary>
        /// Asynchronously releases resources used by the document.
        /// </summary>
        public async ValueTask DisposeAsync() {
            if (_disposed) {
                return;
            }

            Exception? persistenceFailure = null;

            try {
                if (this._spreadSheetDocument != null) {
                    try {
                        if (_persistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                            bool shouldSave = this._spreadSheetDocument.FileOpenAccess != FileAccess.Read &&
                                              (IsPackageDirty || _copyPackageToSourceOnDispose || _copyPackageToFilePathOnDispose);
                            _copyPackageToSourceOnDispose = false;
                            _copyPackageToFilePathOnDispose = false;
                            if (shouldSave) {
                                await SaveAsync().ConfigureAwait(false);
                            }
                        }

                        this._spreadSheetDocument.Dispose();
                    } catch (Exception ex) {
                        persistenceFailure = ex;
                    } finally {
                        this._spreadSheetDocument = null!;
                    }
                }

                try {
                    PersistPackageToSourceIfNeeded();
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }
            } finally {
                if (_ownedOpenStream != null) {
                    try {
                        _ownedOpenStream.Dispose();
                    } catch (Exception ex) {
                        persistenceFailure ??= ex;
                    }
                    _ownedOpenStream = null;
                }

                _lock?.Dispose();
                _disposed = true;
                GC.SuppressFinalize(this);
            }

            if (persistenceFailure != null) {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(persistenceFailure).Throw();
            }
        }

        private void PersistPackageToSourceIfNeeded() {
            if (_packageStream == null) {
                return;
            }

            try {
                if (_copyPackageToSourceOnDispose && _sourceStream != null) {
                    PersistPackageToSource();
                } else if (_copyPackageToFilePathOnDispose && !string.IsNullOrEmpty(FilePath)) {
                    PersistPackageToFilePath();
                }
            } finally {
                DisposeStream(_packageStream);

                if (_copyPackageToSourceOnDispose && _sourceStream != null) {
                    if (!_leaveSourceStreamOpen) {
                        try {
                            _sourceStream.Dispose();
                        } catch {
                            // ignored
                        }
                    } else if (_sourceStream.CanSeek) {
                        try {
                            _sourceStream.Seek(0, SeekOrigin.Begin);
                        } catch {
                            // ignored
                        }
                    }
                }

                _packageStream = null;
                _sourceStream = null;
                _copyPackageToSourceOnDispose = false;
                _copyPackageToFilePathOnDispose = false;
                _leaveSourceStreamOpen = true;
            }
        }

        private void PersistPackageToSource() {
            var packageStream = _packageStream ?? throw new InvalidOperationException("Package stream is not available.");
            var targetStream = _sourceStream ?? throw new InvalidOperationException("Source stream is not available.");

            if (!targetStream.CanSeek) {
                throw new InvalidOperationException("The provided stream must support seeking when SaveOnDispose is enabled.");
            }

            if (packageStream.CanSeek) {
                packageStream.Seek(0, SeekOrigin.Begin);
            }

            targetStream.Seek(0, SeekOrigin.Begin);
            targetStream.SetLength(0);
            packageStream.CopyTo(targetStream, StreamCopyBufferSize);
            targetStream.Flush();
            targetStream.Seek(0, SeekOrigin.Begin);
        }

        private void PersistPackageToFilePath() {
            var packageStream = _packageStream ?? throw new InvalidOperationException("Package stream is not available.");
            if (string.IsNullOrEmpty(FilePath)) {
                throw new InvalidOperationException("File path is not available.");
            }

            if (packageStream.CanSeek) {
                packageStream.Seek(0, SeekOrigin.Begin);
            }

            using var targetStream = new FileStream(FilePath, FileMode.Create, FileAccess.Write, FileShare.None);
            packageStream.CopyTo(targetStream, StreamCopyBufferSize);
            targetStream.Flush();
        }
    }
}
