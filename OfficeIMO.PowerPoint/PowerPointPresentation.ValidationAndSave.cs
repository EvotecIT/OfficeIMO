using OfficeIMO.Drawing.Internal;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>Opens the associated presentation in the operating system's registered application.</summary>
        public void OpenInApplication(string? filePath = null) {
            string target = string.IsNullOrEmpty(filePath) ? _filePath : filePath!;
            if (string.IsNullOrEmpty(target)) {
                throw new InvalidOperationException("The presentation has no associated file path.");
            }
            OfficeFileLauncher.Open(target);
        }

        /// <summary>
        ///     Indicates whether the presentation passes Open XML validation.
        /// </summary>
        public bool DocumentIsValid {
            get {
                if (DocumentValidationErrors.Count > 0) {
                    return false;
                }

                return true;
            }
        }

        /// <summary>
        ///     Gets the list of validation errors for the presentation.
        /// </summary>
        public List<ValidationErrorInfo> DocumentValidationErrors {
            get {
                return ValidateDocument();
            }
        }

        /// <summary>
        ///     Validates the presentation using the specified file format version.
        /// </summary>
        /// <param name="fileFormatVersions">File format version to validate against.</param>
        /// <returns>List of validation errors.</returns>
        /// <example>
        /// <code>
        /// using (var presentation = PowerPointPresentation.Create("test.pptx")) {
        ///     var errors = presentation.ValidateDocument();
        ///     if (errors.Count > 0) {
        ///         // Handle validation errors
        ///     }
        /// }
        /// </code>
        /// </example>
        public List<ValidationErrorInfo> ValidateDocument(FileFormatVersions fileFormatVersions = FileFormatVersions.Microsoft365) {
            ThrowIfDisposed();
            List<ValidationErrorInfo> listErrors = new List<ValidationErrorInfo>();
            OpenXmlValidator validator = new OpenXmlValidator(fileFormatVersions);
            foreach (ValidationErrorInfo error in validator.Validate(_document!)) {
                listErrors.Add(error);
            }

            return listErrors;
        }

        /// <summary>
        ///     Saves all pending changes to the associated file or stream.
        /// </summary>
        public void Save() {
            if (!string.IsNullOrEmpty(_filePath)) {
                Save(_filePath);
                return;
            }
            if (_sourceStream != null) {
                if (IsLegacyBinaryFormat(SourceFormat)) {
                    Save(_sourceStream, SourceFormat);
                } else {
                    Save(_sourceStream);
                }
                return;
            }
            throw new InvalidOperationException(
                "The presentation has no associated destination. Use Save(string) or Save(Stream).");
        }

        /// <summary>Saves the presentation to a file and associates that path with subsequent <see cref="Save()"/> calls.</summary>
        public void Save(string filePath) => Save(filePath, options: null);

        /// <summary>Saves to a file with an explicit conversion-loss policy.</summary>
        public void Save(string filePath, PowerPointSaveOptions? options) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            EnsureDestinationFileWritable(filePath);
            byte[] packageBytes = CreateBytesForPath(filePath, options);
            OfficeFileCommit.WriteAllBytes(filePath, packageBytes);
            _filePath = filePath;
            _discardChangesOnDispose = false;
        }

        /// <summary>
        ///     Saves the presentation to the provided stream without changing its associated destination.
        /// </summary>
        public void Save(Stream destination) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            OfficeStreamWriter.WriteAllBytes(destination, CreatePackageBytesForSave());
            _discardChangesOnDispose = false;
        }

        /// <summary>Saves an independent copy without changing the presentation's associated destination.</summary>
        public void SaveCopy(string filePath) => SaveCopy(filePath, options: null);

        /// <summary>Saves an independent copy with an explicit conversion-loss policy.</summary>
        public void SaveCopy(string filePath, PowerPointSaveOptions? options) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            EnsureDestinationFileWritable(filePath);
            OfficeFileCommit.WriteAllBytes(filePath, CreateBytesForPath(filePath, options));
            _discardChangesOnDispose = false;
        }

        /// <summary>Asynchronously saves an independent copy without changing the presentation's associated destination.</summary>
        public Task SaveCopyAsync(string filePath, CancellationToken cancellationToken = default) =>
            SaveCopyAsync(filePath, options: null, cancellationToken);

        /// <summary>Asynchronously saves an independent copy with an explicit conversion-loss policy.</summary>
        public async Task SaveCopyAsync(string filePath, PowerPointSaveOptions? options,
            CancellationToken cancellationToken = default) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            EnsureDestinationFileWritable(filePath);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] packageBytes = CreateBytesForPath(filePath, options);
            await OfficeFileCommit.WriteAllBytesAsync(filePath, packageBytes, cancellationToken: cancellationToken)
                .ConfigureAwait(false);
            _discardChangesOnDispose = false;
        }

        /// <summary>Encodes the presentation as a PPTX package.</summary>
        public byte[] ToBytes() => CreatePackageBytesForSave();

        /// <summary>Encodes the presentation in the requested PowerPoint format.</summary>
        public byte[] ToBytes(PowerPointFileFormat format, PowerPointSaveOptions? options = null) =>
            IsLegacyBinaryFormat(format) ? CreateLegacyPptBytesForSave(options) : CreatePackageBytesForSave();

        /// <summary>Encodes the presentation in a new writable memory stream positioned at the beginning.</summary>
        public MemoryStream ToStream() => new MemoryStream(ToBytes());

        /// <summary>Encodes the presentation in a new stream using the requested format.</summary>
        public MemoryStream ToStream(PowerPointFileFormat format, PowerPointSaveOptions? options = null) =>
            new MemoryStream(ToBytes(format, options));

        /// <summary>Saves to a stream using an explicit physical PowerPoint format.</summary>
        public void Save(Stream destination, PowerPointFileFormat format, PowerPointSaveOptions? options = null) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            OfficeStreamWriter.WriteAllBytes(destination, ToBytes(format, options));
            _discardChangesOnDispose = false;
        }

        /// <summary>Asynchronously saves to the associated file or stream.</summary>
        public Task SaveAsync(CancellationToken cancellationToken = default) {
            if (!string.IsNullOrEmpty(_filePath)) {
                return SaveAsync(_filePath, cancellationToken);
            }
            if (_sourceStream != null) {
                return IsLegacyBinaryFormat(SourceFormat)
                    ? SaveAsync(_sourceStream, SourceFormat, options: null, cancellationToken)
                    : SaveAsync(_sourceStream, cancellationToken);
            }
            throw new InvalidOperationException(
                "The presentation has no associated destination. Use SaveAsync(string) or SaveAsync(Stream).");
        }

        /// <summary>Asynchronously saves to a file and associates it with subsequent saves.</summary>
        public Task SaveAsync(
            string filePath,
            CancellationToken cancellationToken = default) => SaveAsync(filePath, options: null, cancellationToken);

        /// <summary>Asynchronously saves to a file with an explicit conversion-loss policy.</summary>
        public async Task SaveAsync(
            string filePath,
            PowerPointSaveOptions? options,
            CancellationToken cancellationToken = default) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            EnsureDestinationFileWritable(filePath);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] packageBytes = CreateBytesForPath(filePath, options);
            await OfficeFileCommit.WriteAllBytesAsync(filePath, packageBytes,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            _filePath = filePath;
            _discardChangesOnDispose = false;
        }

        /// <summary>Asynchronously saves once to a caller-owned writable stream without changing the associated destination.</summary>
        public async Task SaveAsync(Stream destination, CancellationToken cancellationToken = default) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            cancellationToken.ThrowIfCancellationRequested();
            byte[] packageBytes = CreatePackageBytesForSave();
            await OfficeStreamWriter.WriteAllBytesAsync(destination, packageBytes, cancellationToken)
                .ConfigureAwait(false);
            _discardChangesOnDispose = false;
        }

        /// <summary>Asynchronously saves once to a stream using an explicit physical PowerPoint format.</summary>
        public async Task SaveAsync(Stream destination, PowerPointFileFormat format,
            PowerPointSaveOptions? options = null, CancellationToken cancellationToken = default) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = ToBytes(format, options);
            await OfficeStreamWriter.WriteAllBytesAsync(destination, bytes, cancellationToken).ConfigureAwait(false);
            _discardChangesOnDispose = false;
        }

        private byte[] CreatePackageBytesForSave() {
            ThrowIfDisposed();
            if (AccessMode == DocumentAccessMode.ReadOnly) {
                throw new InvalidOperationException("The presentation is read-only and cannot be saved.");
            }
            ApplySignatureMutationPolicy();
            foreach (PowerPointSlide slide in _slides) {
                slide.Save();
            }
            PowerPointUtils.UpdateDocumentProperties(_presentationPart);
            PresentationRoot.Save();
            _document!.Save();

            using var packageStream = new MemoryStream();
            using (var clone = _document.Clone(packageStream)) {
                // Dispose finalizes the cloned package before its bytes are committed.
            }
            return packageStream.ToArray();
        }

        private byte[] CreateBytesForPath(string filePath, PowerPointSaveOptions? options) {
            PowerPointFileFormat format = PowerPointPresentationLoadRouting.GetFormat(filePath);
            return IsLegacyBinaryFormat(format)
                ? CreateLegacyPptBytesForSave(options)
                : CreatePackageBytesForSave();
        }

        private byte[] CreateLegacyPptBytesForSave(
            PowerPointSaveOptions? options) {
            if (_legacyPptPackage?.WasEncryptedSource != true) {
                return CreatePlainLegacyPptBytesForSave(options);
            }
            if (options == null
                && TryCopyOriginalEncryptedLegacyPackage(
                    out byte[] originalEncryptedBytes)) {
                return originalEncryptedBytes;
            }
            ApplySignatureMutationPolicy();
            byte[] plainBytes = CreatePlainLegacyPptBytesForSave(options);
            string password = _legacyPptPackage.EncryptionPassword
                ?? throw new InvalidOperationException(
                    "The encrypted binary source password is unavailable. Use SaveEncrypted to persist the presentation.");
            int keySizeBits = options?.LegacyPptEncryptionKeySizeBits
                ?? _legacyPptPackage.EncryptionKeySizeBits ?? 128;
            bool encryptDocumentProperties = options == null
                ? _legacyPptPackage.EncryptedDocumentProperties ?? true
                : options.LegacyPptEncryptDocumentProperties;
            return LegacyPptRc4CryptoApi.EncryptPackage(plainBytes,
                password, keySizeBits, encryptDocumentProperties);
        }

        private byte[] CreatePlainLegacyPptBytesForSave(
            PowerPointSaveOptions? options) {
            ThrowIfDisposed();
            if (AccessMode == DocumentAccessMode.ReadOnly) {
                throw new InvalidOperationException("The presentation is read-only and cannot be saved.");
            }
            if (LastSignatureReport?.Action
                    != PowerPointSignatureMutationAction.Removed
                && TryCopyOriginalLegacyPackage(out byte[] originalBytes)) {
                return originalBytes;
            }
            ApplySignatureMutationPolicy();
            foreach (PowerPointSlide slide in _slides) slide.Save();
            PresentationRoot.Save();
            _document!.Save();
            if (LegacyPptPreservingWriter.TryWritePresentation(this, out byte[] preservedBytes)) {
                return preservedBytes;
            }
            if (LastSignatureReport?.Action == PowerPointSignatureMutationAction.Preserved
                && LastSignatureReport.HasSignatureMetadata) {
                throw new NotSupportedException(
                    "The requested binary rewrite cannot preserve the presentation's digital-signature carrier. " +
                    "Choose RemoveInvalidatedSignatures or make only preservation-aware edits.");
            }
            return LegacyPptWriter.WritePresentation(this, options);
        }

        /// <summary>
        ///     Exports a single slide as a standalone one-slide PowerPoint presentation.
        /// </summary>
        /// <param name="slideIndex">Zero-based index of the slide to export.</param>
        /// <param name="filePath">Destination .pptx path.</param>
        public void ExportSlide(int slideIndex, string filePath) {
            ThrowIfDisposed();
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));

            ValidateSlideIndex(slideIndex);
            using var stream = new MemoryStream();
            ExportSlide(slideIndex, stream);
            OfficeFileCommit.WriteAllBytes(filePath, stream.ToArray());
        }

        /// <summary>
        ///     Exports a single slide as a standalone one-slide PowerPoint presentation.
        /// </summary>
        /// <param name="slideIndex">Zero-based index of the slide to export.</param>
        /// <param name="destination">Writable destination stream.</param>
        public void ExportSlide(int slideIndex, Stream destination) {
            ThrowIfDisposed();
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));

            ValidateSlideIndex(slideIndex);
            using PowerPointPresentation exported = Create();
            exported.ImportSlideForExport(this, slideIndex);
            exported.Save(destination);
        }

        /// <summary>
        /// Saves the presentation with password-to-open encryption. PPT/POT/PPS paths use
        /// legacy RC4 CryptoAPI for compatibility; Open XML paths use modern package encryption.
        /// </summary>
        /// <param name="filePath">Destination path for the encrypted presentation.</param>
        /// <param name="password">Password used to encrypt the presentation package.</param>
        public void SaveEncrypted(string filePath, string password) =>
            SaveEncrypted(filePath, password, options: null);

        /// <summary>
        /// Saves the presentation with password-to-open encryption and explicit
        /// binary conversion or RC4 settings.
        /// </summary>
        /// <param name="filePath">Destination path for the encrypted presentation.</param>
        /// <param name="password">Password used to encrypt the presentation package.</param>
        /// <param name="options">Binary conversion and RC4 key-size options.</param>
        public void SaveEncrypted(string filePath, string password,
            PowerPointSaveOptions? options) {
            ThrowIfDisposed();
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (filePath.Length == 0) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            EnsureDestinationFileWritable(filePath);

            PowerPointFileFormat format = PowerPointPresentationLoadRouting
                .GetFormat(filePath);
            byte[] encryptedBytes = IsLegacyBinaryFormat(format)
                ? CreateEncryptedLegacyPptBytes(password, options)
                : OfficeEncryption.EncryptPackage(CreatePackageBytesForSave(),
                    password);
            OfficeFileCommit.WriteAllBytes(filePath, encryptedBytes);
        }

        /// <summary>
        /// Saves the presentation with password-to-open encryption to a stream. A presentation
        /// loaded from PPT/POT/PPS retains its binary physical format; other presentations use Open XML.
        /// </summary>
        /// <param name="destination">Writable stream receiving the encrypted presentation.</param>
        /// <param name="password">Password used to encrypt the presentation package.</param>
        public void SaveEncrypted(Stream destination, string password) {
            ThrowIfDisposed();
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (password == null) throw new ArgumentNullException(nameof(password));

            byte[] encryptedBytes = IsLegacyBinaryFormat(SourceFormat)
                ? CreateEncryptedLegacyPptBytes(password, options: null)
                : OfficeEncryption.EncryptPackage(CreatePackageBytesForSave(),
                    password);
            OfficeStreamWriter.WriteAllBytes(destination, encryptedBytes);
        }

        /// <summary>Saves a password-encrypted presentation using an explicit physical format.</summary>
        /// <param name="destination">Writable destination stream.</param>
        /// <param name="password">Password used to encrypt the presentation.</param>
        /// <param name="format">Physical PowerPoint format to write.</param>
        /// <param name="options">Binary conversion and RC4 key-size options.</param>
        public void SaveEncrypted(Stream destination, string password,
            PowerPointFileFormat format, PowerPointSaveOptions? options = null) {
            ThrowIfDisposed();
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (password == null) throw new ArgumentNullException(nameof(password));
            byte[] encryptedBytes = IsLegacyBinaryFormat(format)
                ? CreateEncryptedLegacyPptBytes(password, options)
                : OfficeEncryption.EncryptPackage(CreatePackageBytesForSave(),
                    password);
            OfficeStreamWriter.WriteAllBytes(destination, encryptedBytes);
        }

        /// <summary>Encodes a password-encrypted presentation in the requested physical format.</summary>
        /// <param name="password">Password used to encrypt the presentation.</param>
        /// <param name="format">Physical PowerPoint format to encode.</param>
        /// <param name="options">Binary conversion and RC4 key-size options.</param>
        /// <returns>The complete encrypted package bytes.</returns>
        public byte[] ToEncryptedBytes(string password,
            PowerPointFileFormat format = PowerPointFileFormat.Pptx,
            PowerPointSaveOptions? options = null) {
            ThrowIfDisposed();
            if (password == null) throw new ArgumentNullException(nameof(password));
            return IsLegacyBinaryFormat(format)
                ? CreateEncryptedLegacyPptBytes(password, options)
                : OfficeEncryption.EncryptPackage(CreatePackageBytesForSave(),
                    password);
        }

        private byte[] CreateEncryptedLegacyPptBytes(string password,
            PowerPointSaveOptions? options) {
            ApplySignatureMutationPolicy();
            byte[] plainBytes = CreatePlainLegacyPptBytesForSave(options);
            return LegacyPptRc4CryptoApi.EncryptPackage(plainBytes, password,
                options?.LegacyPptEncryptionKeySizeBits ?? 128,
                options?.LegacyPptEncryptDocumentProperties ?? true);
        }

        private static void EnsureDestinationFileWritable(string filePath) {
            if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
            }
        }

    }
}
