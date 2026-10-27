using OfficeIMO.Core.Internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <inheritdoc />
        public void Dispose() {
            if (_disposed) {
                return;
            }

            Exception? pendingException = null;
            PresentationDocument? document = _document;
            try {
                if (document != null && !_discardChangesOnDispose &&
                    _persistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                    bool shouldSave = document.FileOpenAccess != FileAccess.Read;
                    if (shouldSave && _signedPackageOpenFingerprint != null) {
                        shouldSave = !string.Equals(_signedPackageOpenFingerprint,
                            CreatePackageFingerprint(document), StringComparison.Ordinal);
                    }
                    if (shouldSave) {
                        Save();
                    }
                }
            } catch (Exception ex) {
                pendingException = ex;
            } finally {
                try {
                    document?.Dispose();
                } catch (Exception ex) {
                    pendingException ??= ex;
                }
                _document = null;

                try {
                    _packageStream?.Dispose();
                } catch (Exception ex) {
                    pendingException ??= ex;
                }
                _packageStream = null;
                _sourceStream = null;
                _signedPackageOpenFingerprint = null;
                ClearLegacyPptPackageState();
                _discardChangesOnDispose = false;
                _disposed = true;
                GC.SuppressFinalize(this);
            }

            if (pendingException != null) {
                ExceptionDispatchInfo.Capture(pendingException).Throw();
            }
        }

        /// <summary>Asynchronously persists opt-in changes and releases package resources.</summary>
        public async ValueTask DisposeAsync() {
            if (_disposed) {
                return;
            }

            Exception? pendingException = null;
            PresentationDocument? document = _document;
            try {
                if (document != null && !_discardChangesOnDispose &&
                    _persistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                    bool shouldSave = document.FileOpenAccess != FileAccess.Read;
                    if (shouldSave && _signedPackageOpenFingerprint != null) {
                        shouldSave = !string.Equals(_signedPackageOpenFingerprint,
                            CreatePackageFingerprint(document), StringComparison.Ordinal);
                    }
                    if (shouldSave) {
                        await SaveAsync().ConfigureAwait(false);
                    }
                }
            } catch (Exception ex) {
                pendingException = ex;
            } finally {
                try {
                    document?.Dispose();
                } catch (Exception ex) {
                    pendingException ??= ex;
                }
                _document = null;

                try {
                    _packageStream?.Dispose();
                } catch (Exception ex) {
                    pendingException ??= ex;
                }
                _packageStream = null;
                _sourceStream = null;
                _signedPackageOpenFingerprint = null;
                ClearLegacyPptPackageState();
                _discardChangesOnDispose = false;
                _disposed = true;
                GC.SuppressFinalize(this);
            }

            if (pendingException != null) {
                ExceptionDispatchInfo.Capture(pendingException).Throw();
            }
        }

        /// <summary>Creates a detached presentation with explicit persistence.</summary>
        public static PowerPointDocument Create(PowerPointCreateOptions? options = null) =>
            CreateInternal(filePath: null, sourceStream: null, options ?? new PowerPointCreateOptions());

        /// <summary>
        /// Creates a new presentation associated with a file path. The path is not created until
        /// <see cref="Save()"/> is called or SaveOnDispose is explicitly enabled.
        /// </summary>
        public static PowerPointDocument Create(string filePath, PowerPointCreateOptions? options = null) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            return CreateInternal(filePath, sourceStream: null, options ?? new PowerPointCreateOptions());
        }

        /// <summary>
        /// Creates a new presentation associated with a caller-owned stream. The stream is not written until
        /// <see cref="Save()"/> is called or SaveOnDispose is explicitly enabled.
        /// </summary>
        public static PowerPointDocument Create(Stream stream, PowerPointCreateOptions? options = null) {
            OfficeDocumentLifecycle.EnsureAssociatedDestination(stream, nameof(stream));
            PowerPointCreateOptions resolved = options ?? new PowerPointCreateOptions();
            return CreateInternal(filePath: null, stream, resolved);
        }

        private static PowerPointDocument CreateInternal(
            string? filePath,
            Stream? sourceStream,
            PowerPointCreateOptions options) {
            if (options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose &&
                string.IsNullOrEmpty(filePath) && sourceStream == null) {
                throw new ArgumentException("SaveOnDispose requires an associated file path or writable stream.", nameof(options));
            }

            if (options.Template != null) {
                PowerPointDocument copy = Load(options.Template);
                copy._filePath = filePath ?? string.Empty;
                copy._sourceStream = sourceStream;
                copy._persistenceMode = options.PersistenceMode;
                return copy;
            }
            var packageStream = new MemoryStream(StreamBufferSize);
            try {
                PresentationDocument document = PresentationDocument.Create(
                    packageStream, PresentationDocumentType.Presentation, autoSave: false);
                var presentation = new PowerPointDocument(document, filePath ?? string.Empty, isNewDocument: true) {
                    _packageStream = packageStream,
                    _sourceStream = sourceStream,
                    _persistenceMode = options.PersistenceMode
                };
                presentation.PresentationRoot.Save();
                presentation._document?.Save();
                return presentation;
            } catch {
                packageStream.Dispose();
                throw;
            }
        }

        /// <summary>Loads an existing presentation into detached memory.</summary>
        public static PowerPointDocument Load(string filePath,
            PowerPointLoadOptions? options = null) => Load(filePath, options,
                CancellationToken.None);

        /// <summary>Loads an existing presentation into detached memory with cancellation.</summary>
        public static PowerPointDocument Load(string filePath,
            PowerPointLoadOptions? options,
            CancellationToken cancellationToken) {
            if (options?.Password is string password) return AssociateEncryptedDocument(LoadEncrypted(filePath, password, options, cancellationToken), filePath, password);
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            PowerPointLoadOptions resolved = options
                ?? new PowerPointLoadOptions();
            byte[] bytes;
            using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                       FileShare.ReadWrite | FileShare.Delete)) {
                bytes = ReadDocumentInputBytes(source, resolved,
                    cancellationToken, filePath);
            }
            return LoadDocument(bytes, filePath, sourceStream: null,
                resolved, cancellationToken);
        }

        /// <summary>Loads a presentation from a caller-owned stream into memory. Editable writable seekable sources become the associated destination; other sources remain detached.</summary>
        public static PowerPointDocument Load(Stream stream,
            PowerPointLoadOptions? options = null) => Load(stream, options,
                CancellationToken.None);

        /// <summary>Loads a presentation from a caller-owned stream into memory with cancellation.</summary>
        public static PowerPointDocument Load(Stream stream,
            PowerPointLoadOptions? options,
            CancellationToken cancellationToken) {
            if (options?.Password is string password) return AssociateEncryptedDocument(LoadEncrypted(stream, password, options, cancellationToken), null, password, options.AccessMode == DocumentAccessMode.ReadOnly ? null : stream);
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "presentation");
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));

            return LoadDocument(ReadDocumentInputBytes(stream, resolved,
                    cancellationToken),
                filePath: null, stream, resolved, cancellationToken);
        }

        /// <summary>Asynchronously loads an existing presentation into detached memory.</summary>
        public static async Task<PowerPointDocument> LoadAsync(
            string filePath,
            PowerPointLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (options?.Password is string password) return AssociateEncryptedDocument(await LoadEncryptedAsync(filePath, password, options, cancellationToken).ConfigureAwait(false), filePath, password);
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) {
                throw new FileNotFoundException($"File '{fullPath}' doesn't exist.", fullPath);
            }

            PowerPointLoadOptions resolved = options
                ?? new PowerPointLoadOptions();
            using var source = new FileStream(
                fullPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete,
                81920,
                useAsync: true);
            byte[] bytes = await ReadDocumentInputBytesAsync(source,
                resolved, cancellationToken, fullPath)
                .ConfigureAwait(false);
            return LoadDocument(bytes, fullPath, sourceStream: null,
                resolved, cancellationToken);
        }

        /// <summary>Asynchronously loads a presentation from a caller-owned stream.</summary>
        public static async Task<PowerPointDocument> LoadAsync(
            Stream stream,
            PowerPointLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (options?.Password is string password) return AssociateEncryptedDocument(await LoadEncryptedAsync(stream, password, options, cancellationToken).ConfigureAwait(false), null, password, options.AccessMode == DocumentAccessMode.ReadOnly ? null : stream);
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "presentation");
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));
            byte[] bytes = await ReadDocumentInputBytesAsync(stream,
                resolved, cancellationToken)
                .ConfigureAwait(false);
            return LoadDocument(bytes, filePath: null, stream, resolved,
                cancellationToken);
        }

        /// <summary>Loads a password-encrypted presentation into detached memory.</summary>
        internal static PowerPointDocument LoadEncrypted(
            string filePath,
            string password,
            PowerPointLoadOptions? options = null) => LoadEncrypted(filePath,
                password, options, CancellationToken.None);

        /// <summary>Loads a password-encrypted presentation into detached memory with cancellation.</summary>
        internal static PowerPointDocument LoadEncrypted(
            string filePath,
            string password,
            PowerPointLoadOptions? options,
            CancellationToken cancellationToken) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            using var source = new FileStream(filePath, FileMode.Open,
                FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            byte[] encryptedBytes = ReadDocumentInputBytes(source,
                resolved, cancellationToken, filePath);
            ValidatePackageSecurity(encryptedBytes, resolved);
            LegacyBinaryEncryptionKind legacyEncryption =
                PowerPointDocumentLoadRouting
                    .GetLegacyBinaryEncryptionKind(encryptedBytes,
                        resolved.LegacyPptImportOptions);
            if (!resolved.AllowLegacyProjection && legacyEncryption != LegacyBinaryEncryptionKind.NotLegacyBinary)
                throw new NotSupportedException("This load profile accepts Open XML presentations only; binary PPT projection is disabled.");
            if (legacyEncryption == LegacyBinaryEncryptionKind.Encrypted) {
                return LoadEncryptedLegacyPptFromNormalFlow(encryptedBytes,
                    password, PowerPointDocumentLoadRouting.GetFormat(
                        filePath, legacyDefault: true), resolved,
                    cancellationToken);
            }
            ThrowIfUnencryptedLegacyBinary(legacyEncryption);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(
                encryptedBytes, password, cancellationToken,
                ResolvePackageInputLimit(resolved));
            cancellationToken.ThrowIfCancellationRequested();
            return WithSourcePackageInspection(LoadPackage(packageBytes, filePath: null, sourceStream: null, resolved, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
        }

        /// <summary>Asynchronously loads a password-encrypted presentation into detached memory.</summary>
        internal static async Task<PowerPointDocument> LoadEncryptedAsync(
            string filePath,
            string password,
            PowerPointLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) {
                throw new FileNotFoundException($"File '{fullPath}' doesn't exist.", fullPath);
            }
            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            using var source = new FileStream(
                fullPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete,
                81920,
                useAsync: true);
            byte[] encryptedBytes = await ReadDocumentInputBytesAsync(
                source, resolved, cancellationToken, fullPath)
                .ConfigureAwait(false);
            ValidatePackageSecurity(encryptedBytes, resolved);
            LegacyBinaryEncryptionKind legacyEncryption =
                PowerPointDocumentLoadRouting
                    .GetLegacyBinaryEncryptionKind(encryptedBytes,
                        resolved.LegacyPptImportOptions);
            if (!resolved.AllowLegacyProjection && legacyEncryption != LegacyBinaryEncryptionKind.NotLegacyBinary)
                throw new NotSupportedException("This load profile accepts Open XML presentations only; binary PPT projection is disabled.");
            if (legacyEncryption == LegacyBinaryEncryptionKind.Encrypted) {
                return LoadEncryptedLegacyPptFromNormalFlow(encryptedBytes,
                    password, PowerPointDocumentLoadRouting.GetFormat(
                        fullPath, legacyDefault: true), resolved,
                    cancellationToken);
            }
            ThrowIfUnencryptedLegacyBinary(legacyEncryption);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(
                encryptedBytes, password, cancellationToken,
                ResolvePackageInputLimit(resolved));
            cancellationToken.ThrowIfCancellationRequested();
            return WithSourcePackageInspection(LoadPackage(packageBytes, filePath: null, sourceStream: null, resolved, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
        }

        /// <summary>Asynchronously loads a password-encrypted presentation stream into detached memory.</summary>
        internal static async Task<PowerPointDocument> LoadEncryptedAsync(
            Stream stream,
            string password,
            PowerPointLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            byte[] encryptedBytes = await ReadDocumentInputBytesAsync(
                stream, resolved, cancellationToken)
                .ConfigureAwait(false);
            ValidatePackageSecurity(encryptedBytes, resolved);
            LegacyBinaryEncryptionKind legacyEncryption =
                PowerPointDocumentLoadRouting
                    .GetLegacyBinaryEncryptionKind(encryptedBytes,
                        resolved.LegacyPptImportOptions);
            if (!resolved.AllowLegacyProjection && legacyEncryption != LegacyBinaryEncryptionKind.NotLegacyBinary)
                throw new NotSupportedException("This load profile accepts Open XML presentations only; binary PPT projection is disabled.");
            if (legacyEncryption == LegacyBinaryEncryptionKind.Encrypted) {
                return LoadEncryptedLegacyPptFromNormalFlow(encryptedBytes,
                    password, PowerPointFileFormat.Ppt, resolved,
                    cancellationToken);
            }
            ThrowIfUnencryptedLegacyBinary(legacyEncryption);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(
                encryptedBytes, password, cancellationToken,
                ResolvePackageInputLimit(resolved));
            cancellationToken.ThrowIfCancellationRequested();
            return WithSourcePackageInspection(LoadPackage(packageBytes, filePath: null, sourceStream: null, resolved, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
        }

        /// <summary>Loads a password-encrypted presentation stream into detached memory.</summary>
        internal static PowerPointDocument LoadEncrypted(
            Stream stream,
            string password,
            PowerPointLoadOptions? options = null) => LoadEncrypted(stream,
                password, options, CancellationToken.None);

        /// <summary>Loads a password-encrypted presentation stream into detached memory with cancellation.</summary>
        internal static PowerPointDocument LoadEncrypted(
            Stream stream,
            string password,
            PowerPointLoadOptions? options,
            CancellationToken cancellationToken) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            byte[] encryptedBytes = ReadDocumentInputBytes(stream,
                resolved, cancellationToken);
            ValidatePackageSecurity(encryptedBytes, resolved);
            LegacyBinaryEncryptionKind legacyEncryption =
                PowerPointDocumentLoadRouting
                    .GetLegacyBinaryEncryptionKind(encryptedBytes,
                        resolved.LegacyPptImportOptions);
            if (!resolved.AllowLegacyProjection && legacyEncryption != LegacyBinaryEncryptionKind.NotLegacyBinary)
                throw new NotSupportedException("This load profile accepts Open XML presentations only; binary PPT projection is disabled.");
            if (legacyEncryption == LegacyBinaryEncryptionKind.Encrypted) {
                return LoadEncryptedLegacyPptFromNormalFlow(encryptedBytes,
                    password, PowerPointFileFormat.Ppt, resolved,
                    cancellationToken);
            }
            ThrowIfUnencryptedLegacyBinary(legacyEncryption);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(
                encryptedBytes, password, cancellationToken,
                ResolvePackageInputLimit(resolved));
            cancellationToken.ThrowIfCancellationRequested();
            return WithSourcePackageInspection(LoadPackage(packageBytes, filePath: null, sourceStream: null, resolved, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
        }

        private static void ThrowIfUnencryptedLegacyBinary(
            LegacyBinaryEncryptionKind encryptionKind) {
            if (encryptionKind == LegacyBinaryEncryptionKind.Unencrypted) {
                throw new InvalidDataException(
                    "The binary PowerPoint presentation is not password-encrypted. Use PowerPointDocument.Load instead.");
            }
        }

        private static PowerPointDocument LoadPackage(
            byte[] bytes,
            string? filePath,
            Stream? sourceStream,
            PowerPointLoadOptions options,
            bool captureSourcePackageSecurity = true) {
            ValidatePackageSecurity(bytes, options);
            OfficeDocumentLifecycle.Validate(options.AccessMode, options.PersistenceMode, "presentation");
            bool editable = options.AccessMode == DocumentAccessMode.ReadWrite;
            Stream? associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(
                sourceStream,
                options.AccessMode);
            var packageStream = new MemoryStream(bytes.Length + StreamBufferSize);
            packageStream.Write(bytes, 0, bytes.Length);
            packageStream.Position = 0;

            try {
                PresentationDocument document = PresentationDocument.Open(
                    packageStream, editable, CreateOpenSettings(options.OpenSettings));
                bool hasSignature = PowerPointPackageFingerprint
                    .HasSignatureAndEnsurePackageXmlWithinLimit(document);
                var presentation = new PowerPointDocument(document, filePath ?? string.Empty, isNewDocument: false) {
                    _packageStream = packageStream,
                    _sourceStream = associatedStream,
                    _persistenceMode = options.PersistenceMode,
                    _maximumSlideXmlCharacters = options.OpenSettings?.MaxCharactersInPart > 0
                        ? Math.Min(PowerPointXmlReader.MaximumPackageXmlCharacters, options.OpenSettings.MaxCharactersInPart)
                        : PowerPointXmlReader.MaximumPackageXmlCharacters
                };
                if (hasSignature) {
                    presentation._signedPackageOpenFingerprint = CreatePackageFingerprint(document);
                }
                presentation.MarkLoadedFromOpenXml(bytes);
                return captureSourcePackageSecurity ? WithSourcePackageInspection(presentation, bytes, options) : presentation;
            } catch {
                packageStream.Dispose();
                throw;
            }
        }

        private static PowerPointDocument LoadDocument(
            byte[] bytes,
            string? filePath,
            Stream? sourceStream,
            PowerPointLoadOptions options,
            CancellationToken cancellationToken = default) {
            if (PowerPointDocumentLoadRouting.IsLegacyBinary(bytes,
                    filePath)) {
                ValidatePackageSecurity(bytes, options);
                if (!options.AllowLegacyProjection) throw new NotSupportedException("This load profile accepts Open XML presentations only; binary PPT projection is disabled.");
                return WithSourcePackageInspection(LoadLegacyPptFromNormalFlow(bytes, filePath,
                    sourceStream, options, cancellationToken), bytes, options);
            }
            cancellationToken.ThrowIfCancellationRequested();
            return LoadPackage(bytes, filePath, sourceStream, options);
        }

        private static OpenSettings CreateOpenSettings(OfficeOpenXmlLoadSettings? openSettings) =>
            openSettings.ToOpenXml();

        private static void EnsureEncryptedLoadUsesExplicitPersistence(PowerPointLoadOptions options) {
            if (options.PersistenceMode != DocumentPersistenceMode.Explicit) {
                throw new NotSupportedException(
                    "SaveOnDispose is not supported for encrypted PowerPoint sources. Use SaveEncrypted to persist encrypted changes.");
            }
        }

        private static string CreatePackageFingerprint(PresentationDocument document) =>
            PowerPointPackageFingerprint.Create(document);
    }
}
