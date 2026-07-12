using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core;
using OfficeIMO.Core.Internal;
using OfficeIMO.Shared;

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
                _discardChangesOnDispose = false;
                _disposed = true;
                GC.SuppressFinalize(this);
            }

            if (pendingException != null) {
                ExceptionDispatchInfo.Capture(pendingException).Throw();
            }
        }

        /// <summary>Creates a detached presentation with explicit persistence.</summary>
        public static PowerPointPresentation Create() =>
            CreateInternal(filePath: null, sourceStream: null, new PowerPointCreateOptions());

        /// <summary>
        /// Creates a new presentation associated with a file path. The path is not created until
        /// <see cref="Save()"/> is called or SaveOnDispose is explicitly enabled.
        /// </summary>
        public static PowerPointPresentation Create(string filePath, PowerPointCreateOptions? options = null) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            return CreateInternal(filePath, sourceStream: null, options ?? new PowerPointCreateOptions());
        }

        /// <summary>
        /// Creates a new presentation associated with a caller-owned stream. The stream is not written until
        /// <see cref="Save()"/> is called or SaveOnDispose is explicitly enabled.
        /// </summary>
        public static PowerPointPresentation Create(Stream stream, PowerPointCreateOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));
            PowerPointCreateOptions resolved = options ?? new PowerPointCreateOptions();
            if (!OfficeStreamWriter.CanReplaceContents(stream)) {
                throw new ArgumentException("Stream must support seeking when used as an associated destination.", nameof(stream));
            }
            return CreateInternal(filePath: null, stream, resolved);
        }

        private static PowerPointPresentation CreateInternal(
            string? filePath,
            Stream? sourceStream,
            PowerPointCreateOptions options) {
            if (options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose &&
                string.IsNullOrEmpty(filePath) && sourceStream == null) {
                throw new ArgumentException("SaveOnDispose requires an associated file path or writable stream.", nameof(options));
            }

            var packageStream = new MemoryStream(StreamBufferSize);
            try {
                PresentationDocument document = PresentationDocument.Create(
                    packageStream, PresentationDocumentType.Presentation, autoSave: false);
                var presentation = new PowerPointPresentation(document, filePath ?? string.Empty, isNewPresentation: true) {
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
        public static PowerPointPresentation Load(string filePath, PowerPointLoadOptions? options = null) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            byte[] bytes;
            using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                       FileShare.ReadWrite | FileShare.Delete)) {
                using var buffer = new MemoryStream();
                source.CopyTo(buffer);
                bytes = buffer.ToArray();
            }
            return LoadPackage(bytes, filePath, sourceStream: null, options ?? new PowerPointLoadOptions());
        }

        /// <summary>Loads a presentation from a caller-owned stream into memory. Editable writable seekable sources become the associated destination; other sources remain detached.</summary>
        public static PowerPointPresentation Load(Stream stream, PowerPointLoadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            ValidateLifecycle(resolved.AccessMode, resolved.PersistenceMode);
            if (resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && !stream.CanWrite) {
                throw new ArgumentException("Stream must be writable when SaveOnDispose is enabled.", nameof(stream));
            }
            if (resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && !stream.CanSeek) {
                throw new ArgumentException("Stream must support seeking when SaveOnDispose is enabled.", nameof(stream));
            }

            return LoadPackage(ReadAllBytes(stream), filePath: null, stream, resolved);
        }

        /// <summary>Loads a password-encrypted presentation into detached memory.</summary>
        public static PowerPointPresentation LoadEncrypted(
            string filePath,
            string password,
            PowerPointLoadOptions? options = null) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(File.ReadAllBytes(filePath), password);
            return LoadPackage(packageBytes, filePath: null, sourceStream: null, resolved);
        }

        /// <summary>Loads a password-encrypted presentation stream into detached memory.</summary>
        public static PowerPointPresentation LoadEncrypted(
            Stream stream,
            string password,
            PowerPointLoadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(ReadAllBytes(stream), password);
            return LoadPackage(packageBytes, filePath: null, sourceStream: null, resolved);
        }

        private static PowerPointPresentation LoadPackage(
            byte[] bytes,
            string? filePath,
            Stream? sourceStream,
            PowerPointLoadOptions options) {
            ValidateLifecycle(options.AccessMode, options.PersistenceMode);
            bool editable = options.AccessMode == DocumentAccessMode.ReadWrite;
            Stream? associatedStream = editable && sourceStream != null &&
                                       OfficeStreamWriter.CanReplaceContents(sourceStream)
                ? sourceStream
                : null;
            var packageStream = new MemoryStream(bytes.Length + StreamBufferSize);
            packageStream.Write(bytes, 0, bytes.Length);
            packageStream.Position = 0;

            try {
                PresentationDocument document = PresentationDocument.Open(
                    packageStream, editable, CreateOpenSettings(options.OpenSettings));
                var presentation = new PowerPointPresentation(document, filePath ?? string.Empty, isNewPresentation: false) {
                    _packageStream = packageStream,
                    _sourceStream = associatedStream,
                    _persistenceMode = options.PersistenceMode
                };
                if (document.DigitalSignatureOriginPart != null ||
                    document.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null) {
                    presentation._signedPackageOpenFingerprint = CreatePackageFingerprint(document);
                }
                return presentation;
            } catch {
                packageStream.Dispose();
                throw;
            }
        }

        private static OpenSettings CreateOpenSettings(OpenSettings? openSettings) {
            if (openSettings == null) {
                return new OpenSettings { AutoSave = false };
            }
            return new OpenSettings {
                AutoSave = false,
                CompatibilityLevel = openSettings.CompatibilityLevel,
                MarkupCompatibilityProcessSettings = openSettings.MarkupCompatibilityProcessSettings,
                MaxCharactersInPart = openSettings.MaxCharactersInPart
            };
        }

        private static void ValidateLifecycle(
            DocumentAccessMode accessMode,
            DocumentPersistenceMode persistenceMode) {
            if (accessMode == DocumentAccessMode.ReadOnly &&
                persistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                throw new ArgumentException("A read-only presentation cannot use SaveOnDispose persistence.");
            }
        }

        private static void EnsureEncryptedLoadUsesExplicitPersistence(PowerPointLoadOptions options) {
            if (options.PersistenceMode != DocumentPersistenceMode.Explicit) {
                throw new NotSupportedException(
                    "SaveOnDispose is not supported for encrypted PowerPoint sources. Use SaveEncrypted to persist encrypted changes.");
            }
        }

        private static string CreatePackageFingerprint(PresentationDocument document) {
            var parts = new HashSet<OpenXmlPart>();
            foreach (IdPartPair pair in document.Parts) CollectPackageParts(pair.OpenXmlPart, parts);

            var content = new StringBuilder();
            foreach (OpenXmlPart part in parts.OrderBy(item => item.Uri.ToString(), StringComparer.Ordinal)) {
                content.Append(part.Uri).Append('|').Append(part.ContentType).Append('|');
                try {
                    OpenXmlPartRootElement? root = part.RootElement;
                    if (root != null) {
                        content.Append(root.OuterXml);
                    } else {
                        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                        using var memory = new MemoryStream();
                        stream.CopyTo(memory);
                        content.Append(Convert.ToBase64String(memory.ToArray()));
                    }
                } catch (InvalidDataException) {
                    content.Append("unreadable");
                }
                foreach (IdPartPair relationship in part.Parts.OrderBy(item => item.RelationshipId, StringComparer.Ordinal)) {
                    content.Append('|').Append(relationship.RelationshipId).Append('=').Append(relationship.OpenXmlPart.Uri);
                }
            }

            using SHA256 sha = SHA256.Create();
            return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(content.ToString())));
        }

        private static void CollectPackageParts(OpenXmlPart part, ISet<OpenXmlPart> parts) {
            if (!parts.Add(part)) return;
            foreach (IdPartPair child in part.Parts) CollectPackageParts(child.OpenXmlPart, parts);
        }
    }
}
