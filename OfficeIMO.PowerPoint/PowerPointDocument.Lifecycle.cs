using OfficeIMO.Core.Internal;
using System.Runtime.ExceptionServices;
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
            PresentationDocument? document = _sdkDocument;
            try {
                if (!_discardChangesOnDispose &&
                    _persistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                    bool shouldSave = _accessMode != DocumentAccessMode.ReadOnly;
                    // A signed package always opens the SDK at load, where its fingerprint is taken.
                    if (shouldSave && _signedPackageOpenFingerprint != null) {
                        shouldSave = !string.Equals(_signedPackageOpenFingerprint,
                            CreatePackageFingerprint(SdkDocument), StringComparison.Ordinal);
                    }
                    if (shouldSave) {
                        Save();
                    }
                }
            } catch (Exception ex) {
                pendingException = ex;
            } finally {
                try {
                    // A closing editable package rewrites its stream, which nothing reads again.
                    DiscardPendingEdits();
                    OfficePackageLoadSource.AbandonPackageStream(_packageStream);
                    // Over loaded bytes the package holds only managed buffers, so closing it would just recompress
                    // every part that was read before discarding the result.
                    if (_packageStream is not OfficeCopyOnWriteStream) document?.Dispose();
                } catch (Exception ex) {
                    pendingException ??= ex;
                }
                _sdkDocument = null;

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
            PresentationDocument? document = _sdkDocument;
            try {
                if (!_discardChangesOnDispose &&
                    _persistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                    bool shouldSave = _accessMode != DocumentAccessMode.ReadOnly;
                    // A signed package always opens the SDK at load, where its fingerprint is taken.
                    if (shouldSave && _signedPackageOpenFingerprint != null) {
                        shouldSave = !string.Equals(_signedPackageOpenFingerprint,
                            CreatePackageFingerprint(SdkDocument), StringComparison.Ordinal);
                    }
                    if (shouldSave) {
                        await SaveAsync().ConfigureAwait(false);
                    }
                }
            } catch (Exception ex) {
                pendingException = ex;
            } finally {
                try {
                    // A closing editable package rewrites its stream, which nothing reads again.
                    DiscardPendingEdits();
                    OfficePackageLoadSource.AbandonPackageStream(_packageStream);
                    // Over loaded bytes the package holds only managed buffers, so closing it would just recompress
                    // every part that was read before discarding the result.
                    if (_packageStream is not OfficeCopyOnWriteStream) document?.Dispose();
                } catch (Exception ex) {
                    pendingException ??= ex;
                }
                _sdkDocument = null;

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
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
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

        /// <summary>Creates neutral native scaffolding without authoring theme defaults over imported appearance.</summary>
        internal static PowerPointDocument CreateForImport(Stream? sourceStream = null) {
            if (sourceStream != null) OfficeDocumentLifecycle.EnsureAssociatedDestination(sourceStream, nameof(sourceStream));
            return CreateInternal(filePath: null, sourceStream, new PowerPointCreateOptions(), applyTheme: false);
        }

        private static PowerPointDocument CreateInternal(
            string? filePath,
            Stream? sourceStream,
            PowerPointCreateOptions options,
            bool applyTheme = true) {
            if (options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose &&
                string.IsNullOrEmpty(filePath) && sourceStream == null) {
                throw new ArgumentException("SaveOnDispose requires an associated file path or writable stream.", nameof(options));
            }

            if (options.Template != null) {
                PowerPointDocument copy = Load(options.Template);
                copy._filePath = filePath ?? string.Empty;
                copy._sourceStream = sourceStream;
                copy._persistenceMode = options.PersistenceMode;
                if (options.Theme != null) copy.ApplyLayoutTheme(options.Theme);
                return copy;
            }
            OfficeTheme? theme = applyTheme ? options.Theme ?? OfficeTheme.Default : null;
            bool themeApplied = ReferenceEquals(theme, OfficeTheme.Default);
            byte[] packageBytes = themeApplied ? DefaultThemedPackage.Value : DefaultPackage.Value;
            var packageStream = new MemoryStream(packageBytes.Length + StreamBufferSize);
            PresentationDocument? document = null;
            try {
                packageStream.Write(packageBytes, 0, packageBytes.Length);
                packageStream.Position = 0;
                document = PresentationDocument.Open(
                    packageStream, true, new OpenSettings { AutoSave = false });
                var presentation = new PowerPointDocument(document, filePath ?? string.Empty, isNewDocument: false) {
                    _packageStream = packageStream,
                    _sourceStream = sourceStream,
                    _persistenceMode = options.PersistenceMode
                };
                // Package structure is reusable; creation dates and authoring state belong to this instance.
                DateTime timestamp = DateTime.UtcNow;
                document.PackageProperties.Created = timestamp;
                document.PackageProperties.Modified = timestamp;
                if (themeApplied) {
                    presentation._baseTheme = theme;
                    presentation._layoutTheme = PowerPointLayoutTheme.From(theme);
                } else if (theme != null) {
                    presentation.ApplyLayoutTheme(theme);
                }
                presentation.PresentationRoot.Save();
                document.Save();
                return presentation;
            } catch {
                try {
                    document?.Dispose();
                } finally {
                    packageStream.Dispose();
                }
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
            ArgumentNullException.ThrowIfNull(filePath);
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            PowerPointLoadOptions resolved = options
                ?? new PowerPointLoadOptions();
            byte[] bytes;
            using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                       FileShare.ReadWrite | FileShare.Delete)) {
                bytes = PowerPointDocumentInput.ReadDocumentInputBytes(source, resolved,
                    cancellationToken, filePath);
            }
            return LoadDocument(bytes, filePath, sourceStream: null,
                resolved, cancellationToken);
        }

        /// <summary>Loads a presentation from a caller-owned stream. Editable writable seekable sources become the associated destination; other sources remain detached.</summary>
        /// <param name="stream">Readable source. The caller keeps ownership and disposes it.</param>
        /// <param name="options">Access, persistence and package options.</param>
        /// <remarks>
        /// <para>An editable load reads the stream into one in-memory copy, which the package is opened over, and the
        /// stream can be disposed or changed afterwards. Saving writes back to a writable, seekable stream.</para>
        /// <para>A read-only load of a seekable Open XML stream inspects and opens the package directly.
        /// Keep <paramref name="stream"/> open and unchanged until the presentation is disposed, and give the
        /// presentation exclusive use of the stream during reads. Each package read preserves the caller's stream
        /// position. The presentation never writes to or disposes the stream, and cannot be saved. Non-seekable
        /// streams and encrypted or binary sources are buffered and detached from their source. Compatibility
        /// data may require retaining an original package snapshot.</para>
        /// </remarks>
        public static PowerPointDocument Load(Stream stream,
            PowerPointLoadOptions? options = null) => Load(stream, options,
                CancellationToken.None);

        /// <summary>Loads a presentation from a caller-owned stream with cancellation.</summary>
        /// <param name="stream">Readable source. The caller keeps ownership and disposes it.</param>
        /// <param name="options">Access, persistence and package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <remarks>Stream lifetime follows <see cref="Load(Stream, PowerPointLoadOptions?)"/>. The async
        /// in-place path applies to standard <see cref="MemoryStream"/> and <see cref="FileStream"/> sources.
        /// Custom streams are buffered.</remarks>
        public static PowerPointDocument Load(Stream stream,
            PowerPointLoadOptions? options,
            CancellationToken cancellationToken) {
            if (options?.Password is string password) return AssociateEncryptedDocument(LoadEncrypted(stream, password, options, cancellationToken), null, password, options.AccessMode == DocumentAccessMode.ReadOnly ? null : stream);
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "presentation");
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));
            if (CanLoadInPlace(stream, resolved)) {
                return LoadPackageInPlace(stream, resolved, cancellationToken);
            }

            return LoadDocument(PowerPointDocumentInput.ReadDocumentInputBytes(stream, resolved,
                    cancellationToken),
                filePath: null, stream, resolved, cancellationToken);
        }

        /// <summary>Asynchronously loads an existing presentation into detached memory.</summary>
        public static async Task<PowerPointDocument> LoadAsync(
            string filePath,
            PowerPointLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (options?.Password is string password) return AssociateEncryptedDocument(await LoadEncryptedAsync(filePath, password, options, cancellationToken).ConfigureAwait(false), filePath, password);
            ArgumentNullException.ThrowIfNull(filePath);
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
            byte[] bytes = await PowerPointDocumentInput.ReadDocumentInputBytesAsync(source,
                resolved, cancellationToken, fullPath)
                .ConfigureAwait(false);
            return LoadDocument(bytes, fullPath, sourceStream: null,
                resolved, cancellationToken);
        }

        /// <summary>Asynchronously loads a presentation from a caller-owned stream.</summary>
        /// <param name="stream">Readable source. The caller keeps ownership and disposes it.</param>
        /// <param name="options">Access, persistence and package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <remarks>Stream lifetime follows <see cref="Load(Stream, PowerPointLoadOptions?)"/>.</remarks>
        public static async Task<PowerPointDocument> LoadAsync(
            Stream stream,
            PowerPointLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (options?.Password is string password) return AssociateEncryptedDocument(await LoadEncryptedAsync(stream, password, options, cancellationToken).ConfigureAwait(false), null, password, options.AccessMode == DocumentAccessMode.ReadOnly ? null : stream);
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "presentation");
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));
            if (await CanLoadInPlaceAsync(stream, resolved, cancellationToken).ConfigureAwait(false)) {
                return await OfficePackageLoadSource.LoadInPlaceAsync(stream,
                    () => LoadPackageInPlace(stream, resolved, cancellationToken), cancellationToken).ConfigureAwait(false);
            }
            byte[] bytes = await PowerPointDocumentInput.ReadDocumentInputBytesAsync(stream,
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
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentNullException.ThrowIfNull(password);
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            using var source = new FileStream(filePath, FileMode.Open,
                FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            byte[] encryptedBytes = PowerPointDocumentInput.ReadDocumentInputBytes(source,
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
                PowerPointDocumentInput.ResolvePackageInputLimit(resolved));
            cancellationToken.ThrowIfCancellationRequested();
            return WithSourcePackageInspection(LoadPackage(packageBytes, filePath: null, sourceStream: null, resolved, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
        }

        /// <summary>Asynchronously loads a password-encrypted presentation into detached memory.</summary>
        internal static async Task<PowerPointDocument> LoadEncryptedAsync(
            string filePath,
            string password,
            PowerPointLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentNullException.ThrowIfNull(password);
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
            byte[] encryptedBytes = await PowerPointDocumentInput.ReadDocumentInputBytesAsync(
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
                PowerPointDocumentInput.ResolvePackageInputLimit(resolved));
            cancellationToken.ThrowIfCancellationRequested();
            return WithSourcePackageInspection(LoadPackage(packageBytes, filePath: null, sourceStream: null, resolved, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
        }

        /// <summary>Asynchronously loads a password-encrypted presentation stream into detached memory.</summary>
        internal static async Task<PowerPointDocument> LoadEncryptedAsync(
            Stream stream,
            string password,
            PowerPointLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(password);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            byte[] encryptedBytes = await PowerPointDocumentInput.ReadDocumentInputBytesAsync(
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
                PowerPointDocumentInput.ResolvePackageInputLimit(resolved));
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
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(password);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            byte[] encryptedBytes = PowerPointDocumentInput.ReadDocumentInputBytes(stream,
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
                PowerPointDocumentInput.ResolvePackageInputLimit(resolved));
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
            bool captureSourcePackageSecurity = true,
            bool allowDeferredSdk = false) {
            ValidatePackageSecurity(bytes, options);
            OfficeDocumentLifecycle.Validate(options.AccessMode, options.PersistenceMode, "presentation");
            Stream? associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(
                sourceStream,
                options.AccessMode);
            // The package reads the loaded array directly; an editable package copies it only when it first writes.
            Stream packageStream = OfficePackageLoadSource.CreatePackageStream(bytes, options.AccessMode);

            try {
                PowerPointDocument presentation = OpenPackage(packageStream, filePath, associatedStream, options, allowDeferredSdk);
                presentation.MarkLoadedFromOpenXml(bytes);
                return captureSourcePackageSecurity ? WithSourcePackageInspection(presentation, bytes, options) : presentation;
            } catch {
                packageStream.Dispose();
                throw;
            }
        }

        private static bool CanLoadInPlace(Stream stream, PowerPointLoadOptions options) =>
            OfficePackageLoadSource.CanReadInPlace(stream, options.AccessMode, PowerPointDocumentInput.ResolvePackageInputLimit(options));

        private static Task<bool> CanLoadInPlaceAsync(Stream stream, PowerPointLoadOptions options,
            CancellationToken cancellationToken) =>
            OfficePackageLoadSource.CanReadInPlaceAsync(stream, options.AccessMode, PowerPointDocumentInput.ResolvePackageInputLimit(options), cancellationToken);

        /// <summary>Opens a read-only Open XML presentation over the caller's seekable stream without copying it.</summary>
        private static PowerPointDocument LoadPackageInPlace(Stream stream, PowerPointLoadOptions options,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            long originalPosition = stream.Position;
            OfficePackageSourceView packageStream = OfficePackageLoadSource.OpenInPlace(stream);
            try {
                OfficePackageSecurityReport? packageReport = OfficePackageLoadSource.InspectInPlace(packageStream,
                    options.PackageSecurity, options.CaptureSourcePackageSecurity, cancellationToken);
                PowerPointDocument presentation = OpenPackage(packageStream, filePath: null, associatedStream: null, options, allowDeferredSdk: true);
                try {
                    presentation.MarkLoadedFromOpenXml(packageStream, cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                } catch {
                    presentation.Dispose();
                    throw;
                }
                return WithSourcePackageInspection(presentation, sourceBytes: null, options, report: packageReport);
            } catch {
                packageStream.Dispose();
                throw;
            } finally {
                stream.Position = originalPosition;
            }
        }

        private static PowerPointDocument OpenPackage(Stream packageStream, string? filePath,
            Stream? associatedStream, PowerPointLoadOptions options, bool allowDeferredSdk = false) {
            if (allowDeferredSdk && CanDeferSdk(packageStream, options, out int slideCount)) {
                return new PowerPointDocument(filePath ?? string.Empty, options.AccessMode) {
                    _deferredSlideCount = slideCount,
                    _packageStream = packageStream,
                    _sourceStream = associatedStream,
                    _persistenceMode = options.PersistenceMode,
                    // The caller may change its settings after the load; the SDK opens with them as they were.
                    _deferredOpenSettings = options.OpenSettings == null ? null : new OfficeOpenXmlLoadSettings {
                        CompatibilityLevel = options.OpenSettings.CompatibilityLevel,
                        MarkupCompatibilityMode = options.OpenSettings.MarkupCompatibilityMode,
                        MarkupCompatibilityTargetVersion = options.OpenSettings.MarkupCompatibilityTargetVersion,
                        MaxCharactersInPart = options.OpenSettings.MaxCharactersInPart
                    },
                    _maximumSlideXmlCharacters = options.OpenSettings?.MaxCharactersInPart > 0
                        ? Math.Min(PowerPointXmlReader.MaximumPackageXmlCharacters, options.OpenSettings.MaxCharactersInPart)
                        : PowerPointXmlReader.MaximumPackageXmlCharacters
                };
            }
            PresentationDocument document = options.AccessMode == DocumentAccessMode.ReadWrite
                ? PresentationDocument.Open(packageStream, true, CreateOpenSettings(options.OpenSettings))
                : OfficeIMO.OpenXml.Internal.OfficeOpenXmlReadOnlyPackage.OpenPresentation(packageStream, CreateOpenSettings(options.OpenSettings));
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
            return presentation;
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
            return LoadPackage(bytes, filePath, sourceStream, options, allowDeferredSdk: true);
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
