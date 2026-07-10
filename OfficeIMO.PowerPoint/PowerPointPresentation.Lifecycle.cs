using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint.Fluent;
using OfficeIMO.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <inheritdoc />
        public void Dispose() {
            if (_disposed) {
                return;
            }

            Exception? pendingException = null;
            PresentationDocument? document = _document;
            try {
                if (document != null) {
                    if (_copyPackageToSourceOnDispose || _saveOnDispose) {
                        Save();
                    }
                }
            } catch (Exception ex) {
                pendingException = ex;
            } finally {
                try {
                    document?.Dispose();
                } catch (Exception ex) {
                    if (pendingException == null) pendingException = ex;
                }
                _document = null;
                try {
                    PersistPackageToSourceIfNeeded(persistChanges: pendingException == null);
                } catch (Exception ex) when (pendingException == null) {
                    pendingException = ex;
                }
                _saveOnDispose = false;
                _disposed = true;
            }

            if (pendingException != null) {
                ExceptionDispatchInfo.Capture(pendingException).Throw();
            }
        }

        /// <summary>
        ///     Creates a new PowerPoint presentation at the specified file path.
        /// </summary>
        /// <param name="filePath">Path where the presentation file will be created.</param>
        public static PowerPointPresentation Create(string filePath) {
            PresentationDocument document = PresentationDocument.Create(filePath,
                PresentationDocumentType.Presentation, autoSave: false);
            PowerPointPresentation presentation = new(document, filePath, isNewPresentation: true);
            presentation._saveOnDispose = true;
            presentation.PresentationRoot.Save();
            presentation._document?.Save();
            return presentation;
        }

        /// <summary>
        ///     Creates a new PowerPoint presentation in memory and optionally persists it to the provided stream on dispose.
        /// </summary>
        /// <param name="stream">Destination stream for the presentation package.</param>
        /// <param name="autoSave">When true, writes the package back to the stream on dispose.</param>
        public static PowerPointPresentation Create(Stream stream, bool autoSave = true) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));
            if (autoSave && !stream.CanSeek) {
                throw new ArgumentException("Stream must support seeking when autoSave is enabled.", nameof(stream));
            }

            Stream packageStream = autoSave
                ? new NonDisposingMemoryStream(StreamBufferSize)
                : new MemoryStream(StreamBufferSize);

            PresentationDocument document = PresentationDocument.Create(packageStream, PresentationDocumentType.Presentation, autoSave: true);
            PowerPointPresentation presentation = new(document, string.Empty, isNewPresentation: true);
            presentation.PresentationRoot.Save();
            presentation._document?.Save();
            presentation.ConfigureStreamCopy(packageStream, stream, autoSave, leaveSourceStreamOpen: true);
            return presentation;
        }

        /// <summary>
        ///     Opens an existing PowerPoint presentation.
        /// </summary>
        /// <param name="filePath">Path of the presentation file to open.</param>
        public static PowerPointPresentation Open(string filePath) {
            PresentationDocument document = PresentationDocument.Open(filePath, true,
                new OpenSettings { AutoSave = false });
            PowerPointPresentation presentation = new(document, filePath, isNewPresentation: false);
            presentation._saveOnDispose = true;
            return presentation;
        }

        /// <summary>
        ///     Opens an existing PowerPoint presentation in read-only mode (no writes, no repairs).
        /// </summary>
        /// <param name="filePath">Path of the presentation file to open.</param>
        public static PowerPointPresentation OpenRead(string filePath) {
            PresentationDocument document = PresentationDocument.Open(filePath, false);
            return new PowerPointPresentation(document, filePath, isNewPresentation: false);
        }

        /// <summary>
        ///     Opens a password-encrypted Office Open XML PowerPoint presentation.
        /// </summary>
        /// <param name="filePath">Path of the encrypted presentation file to open.</param>
        /// <param name="password">Password used to decrypt the presentation package.</param>
        /// <param name="readOnly">Open the decrypted package in read-only mode.</param>
        public static PowerPointPresentation OpenEncrypted(string filePath, string password, bool readOnly = false) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            byte[] encryptedBytes = File.ReadAllBytes(filePath);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            var packageStream = new NonDisposingMemoryStream(packageBytes.Length + StreamBufferSize);
            packageStream.Write(packageBytes, 0, packageBytes.Length);
            packageStream.Position = 0;

            PresentationDocument document = PresentationDocument.Open(packageStream, !readOnly);
            PowerPointPresentation presentation = new(document, filePath, isNewPresentation: false);
            presentation.ConfigureStreamCopy(packageStream, null, copyPackageToSourceOnDispose: false, leaveSourceStreamOpen: true);
            return presentation;
        }

        /// <summary>
        ///     Opens a PowerPoint presentation from a stream.
        /// </summary>
        /// <param name="stream">Source stream containing the presentation package.</param>
        /// <param name="readOnly">Open the document in read-only mode.</param>
        /// <param name="autoSave">When true, writes the package back to the stream on dispose.</param>
        public static PowerPointPresentation Open(Stream stream, bool readOnly = false, bool autoSave = false) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            bool shouldCopyBack = autoSave && !readOnly;
            if (shouldCopyBack) {
                if (!stream.CanWrite) {
                    throw new ArgumentException("Stream must be writable when autoSave is enabled for editable documents.", nameof(stream));
                }
                if (!stream.CanSeek) {
                    throw new ArgumentException("Stream must support seeking when autoSave is enabled for editable documents.", nameof(stream));
                }
            }

            var bytes = ReadAllBytes(stream);
            Stream packageStream = shouldCopyBack
                ? new NonDisposingMemoryStream(bytes.Length + StreamBufferSize)
                : new MemoryStream(bytes.Length + StreamBufferSize);
            packageStream.Write(bytes, 0, bytes.Length);
            packageStream.Position = 0;

            PresentationDocument document = PresentationDocument.Open(packageStream, !readOnly);
            PowerPointPresentation presentation = new(document, string.Empty, isNewPresentation: false);
            presentation.ConfigureStreamCopy(packageStream, stream, shouldCopyBack, leaveSourceStreamOpen: true);
            return presentation;
        }

        /// <summary>
        ///     Opens a password-encrypted Office Open XML PowerPoint presentation from a stream.
        /// </summary>
        /// <param name="stream">Source stream containing the encrypted presentation package.</param>
        /// <param name="password">Password used to decrypt the presentation package.</param>
        /// <param name="readOnly">Open the decrypted package in read-only mode.</param>
        public static PowerPointPresentation OpenEncrypted(Stream stream, string password, bool readOnly = false) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            byte[] encryptedBytes = ReadAllBytes(stream);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            var packageStream = new NonDisposingMemoryStream(packageBytes.Length + StreamBufferSize);
            packageStream.Write(packageBytes, 0, packageBytes.Length);
            packageStream.Position = 0;

            PresentationDocument document = PresentationDocument.Open(packageStream, !readOnly);
            PowerPointPresentation presentation = new(document, string.Empty, isNewPresentation: false);
            presentation.ConfigureStreamCopy(packageStream, null, copyPackageToSourceOnDispose: false, leaveSourceStreamOpen: true);
            return presentation;
        }

        private void PersistPackageToSourceIfNeeded(bool persistChanges) {
            if (_packageStream == null) {
                return;
            }

            try {
                if (persistChanges && _copyPackageToSourceOnDispose && _sourceStream != null) {
                    PersistPackageToSource();
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
                _leaveSourceStreamOpen = true;
            }
        }

        private void PersistPackageToSource() {
            var packageStream = _packageStream ?? throw new InvalidOperationException("Package stream is not available.");
            var targetStream = _sourceStream ?? throw new InvalidOperationException("Source stream is not available.");

            if (!targetStream.CanSeek) {
                throw new InvalidOperationException("The provided stream must support seeking when autoSave is enabled.");
            }

            if (packageStream.CanSeek) {
                packageStream.Seek(0, SeekOrigin.Begin);
            }

            targetStream.Seek(0, SeekOrigin.Begin);
            targetStream.SetLength(0);
            packageStream.CopyTo(targetStream);
            targetStream.Flush();
            targetStream.Seek(0, SeekOrigin.Begin);
        }

        private static void DisposeStream(Stream stream) {
            if (stream is NonDisposingMemoryStream ndms) {
                ndms.DisposeUnderlying();
            } else {
                stream.Dispose();
            }
        }

    }
}
