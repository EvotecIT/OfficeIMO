using OfficeIMO.Drawing.Internal;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Write;
using OfficeIMO.Word.Fluent;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides functionality for creating, loading and manipulating Word documents.
    /// </summary>
    public partial class WordDocument : IDisposable {

        /// <summary>Opens the associated document in the operating system's registered application.</summary>
        public void OpenInApplication(string? filePath = null) {
            string? target = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
            if (string.IsNullOrEmpty(target)) {
                throw new InvalidOperationException("The document has no associated file path.");
            }
            OfficeFileLauncher.Open(target!);
        }

        /// <summary>
        /// Copies package properties. Clone and SaveAs don't actually clone document properties for some reason, so they must be copied manually
        /// </summary>
        /// <param name="src"></param>
        /// <param name="dest"></param>
        // IPackageProperties is currently marked as experimental (OOXML0001).
        // There is no non-experimental alternative available yet.
#pragma warning disable OOXML0001
        private static void CopyPackageProperties(IPackageProperties src, IPackageProperties dest) {
            dest.Category = src.Category;
            dest.ContentStatus = src.ContentStatus;
            dest.ContentType = src.ContentType;
            dest.Created = src.Created;
            dest.Creator = src.Creator;
            dest.Description = src.Description;
            dest.Identifier = src.Identifier;
            dest.Keywords = src.Keywords;
            dest.Language = src.Language;
            dest.LastModifiedBy = src.LastModifiedBy;
            dest.LastPrinted = src.LastPrinted;
            dest.Modified = src.Modified;
            dest.Revision = src.Revision;
            dest.Subject = src.Subject;
            dest.Title = src.Title;
            dest.Version = src.Version;
        }
#pragma warning restore OOXML0001

        private void SaveFileCore(string? filePath, WordSaveOptions? options) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }
            EnsureSignedDocumentSaveAllowed(options, "Save");
            PreSaving();

            if (this._wordprocessingDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            string? resolvedPath = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
            string target = resolvedPath
                ?? throw new InvalidOperationException("This document is not associated with a file path. Provide a file path or save to a writable stream.");
            if (IsLegacyDocPath(target)) {
                SaveLegacyDocFile(target, options: options);
                return;
            }

            SaveOpenXmlFile(target, updateFilePath: true, options);
        }

        /// <summary>
        /// Saves the document as a password-encrypted Office Open XML package.
        /// </summary>
        /// <param name="filePath">Destination path.</param>
        /// <param name="password">Password used to encrypt the document package.</param>
        /// <param name="saveOptions">Optional save policy settings.</param>
        public void SaveEncrypted(string filePath, string password, WordSaveOptions? saveOptions = null) {
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            }
            string target = filePath;
            EnsureDestinationFileWritable(target);

            byte[] packageBytes = ToBytes(WordFileFormat.Docx, saveOptions);
            byte[] encryptedBytes = OfficeEncryption.EncryptPackage(packageBytes, password);
            OfficeFileCommit.WriteAllBytes(target, encryptedBytes);
            FilePath = target;
        }

        /// <summary>
        /// Saves the document as a password-encrypted Office Open XML package to a stream.
        /// </summary>
        /// <param name="destination">Writable stream receiving the encrypted document.</param>
        /// <param name="password">Password used to encrypt the document package.</param>
        public void SaveEncrypted(Stream destination, string password) {
            SaveEncrypted(destination, password, saveOptions: null);
        }

        /// <summary>
        /// Saves the document as a password-encrypted Office Open XML package to a stream.
        /// </summary>
        /// <param name="destination">Writable stream receiving the encrypted document.</param>
        /// <param name="password">Password used to encrypt the document package.</param>
        /// <param name="saveOptions">Optional save policy settings.</param>
        public void SaveEncrypted(Stream destination, string password, WordSaveOptions? saveOptions) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (password == null) throw new ArgumentNullException(nameof(password));
            byte[] packageBytes = ToBytes(WordFileFormat.Docx, saveOptions);
            OfficeEncryption.EncryptPackageToStream(packageBytes, password, destination);
        }

        /// <summary>
        /// Save WordDocument to where it was open from
        /// </summary>
        public void Save() {
            Save(options: null);
        }

        /// <summary>Saves to the associated destination with optional save settings.</summary>
        public void Save(WordSaveOptions? options) {
            if (string.IsNullOrEmpty(FilePath) && OriginalStream != null) {
                Save(OriginalStream, options);
            } else {
                SaveFileCore(FilePath, options);
            }
        }

        /// <summary>
        /// Save WordDocument to given filePath
        /// </summary>
        /// <param name="filePath"></param>
        public void Save(string filePath) {
            EnsureExplicitFilePath(filePath);
            SaveFileCore(filePath, options: null);
        }

        /// <summary>
        /// Save WordDocument to the given file path with optional save policy settings.
        /// </summary>
        /// <param name="filePath">Destination path.</param>
        /// <param name="options">Optional save policy settings.</param>
        public void Save(string filePath, WordSaveOptions? options) {
            EnsureExplicitFilePath(filePath);
            SaveFileCore(filePath, options);
        }

        // Note: Save() already normalizes table grids for consistent viewing across
        // Word Online/Google Docs without changing authoring semantics. No extra
        // save variants are needed.

        private void SaveCopyCore(string filePath, WordSaveOptions? options) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }
            if (string.IsNullOrEmpty(filePath)) {
                throw new ArgumentException("File path cannot be empty", nameof(filePath));
            }

            EnsureSignedDocumentSaveAllowed(options, "SaveCopy");
            PreSaving();

            if (_wordprocessingDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            if (IsLegacyDocPath(filePath)) {
                SaveLegacyDocFile(filePath, updateFilePath: false, options);
                return;
            }

            SaveOpenXmlFile(filePath, updateFilePath: false, options);
        }

        /// <summary>Saves an independent copy without changing this document's associated destination.</summary>
        /// <param name="filePath">Destination path.</param>
        /// <param name="options">Optional save policy settings.</param>
        public void SaveCopy(string filePath, WordSaveOptions? options = null) {
            SaveCopyCore(filePath, options);
        }

        /// <summary>Asynchronously saves an independent copy without changing this document's associated destination.</summary>
        public async Task SaveCopyAsync(
            string filePath,
            WordSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("File path cannot be empty", nameof(filePath));
            }

            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }

            EnsureSignedDocumentSaveAllowed(options, "SaveCopyAsync");
            PreSaving();
            if (_wordprocessingDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = CreatePathBytesAfterPreflight(filePath, options);
            await OfficeFileCommit.WriteAllBytesAsync(filePath, bytes, cancellationToken: cancellationToken)
                .ConfigureAwait(false);
        }

        /// <summary>Encodes the document in the selected physical format.</summary>
        /// <param name="format">DOCX or legacy DOC output.</param>
        /// <param name="options">Optional save policy settings.</param>
        public byte[] ToBytes(WordFileFormat format = WordFileFormat.Docx, WordSaveOptions? options = null) =>
            ToWordBytes(format, options);

        /// <summary>Encodes the document in a new writable memory stream positioned at the beginning.</summary>
        public MemoryStream ToStream(WordFileFormat format = WordFileFormat.Docx, WordSaveOptions? options = null) =>
            new MemoryStream(ToBytes(format, options));

        private byte[] ToWordBytes(WordFileFormat format, WordSaveOptions? options) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }

            EnsureSignedDocumentSaveAllowed(options, "ToBytes");
            PreSaving();

            if (_wordprocessingDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            if (format == WordFileFormat.Doc) {
                return LegacyDocWriter.WriteDocument(this, options);
            }

            EnsureLegacyDocSaveDoesNotDropImportedContent(options);

            _wordprocessingDocument.Save();
            return CreateOpenXmlBytesAfterSave();
        }

        /// <summary>
        /// Asynchronously saves the document.
        /// </summary>
        /// <param name="filePath">Optional path to save to.</param>
        /// <param name="options">Optional save policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        private async Task SaveFileAsyncCore(string? filePath, WordSaveOptions? options, CancellationToken cancellationToken) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }

            if (string.IsNullOrEmpty(filePath) && string.IsNullOrEmpty(FilePath) && OriginalStream != null) {
                await SaveAsync(OriginalStream, WordFileFormat.Docx, options, cancellationToken).ConfigureAwait(false);
                return;
            }

            EnsureSignedDocumentSaveAllowed(options, "SaveAsync");
            PreSaving();

            if (this._wordprocessingDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            string? resolvedPath = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
            string target = resolvedPath
                ?? throw new InvalidOperationException("This document is not associated with a file path. Provide a file path or save to a writable stream.");
            byte[] bytes = CreatePathBytesAfterPreflight(target, options);
            await OfficeFileCommit.WriteAllBytesAsync(target, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
            FilePath = target;
        }

        /// <summary>
        /// Asynchronously saves the document to where it was open from.
        /// </summary>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(CancellationToken cancellationToken = default) {
            return SaveAsync(options: null, cancellationToken);
        }

        /// <summary>Asynchronously saves to the associated destination with optional save settings.</summary>
        public Task SaveAsync(WordSaveOptions? options, CancellationToken cancellationToken = default) {
            if (string.IsNullOrEmpty(FilePath) && OriginalStream != null) {
                return SaveAsync(OriginalStream, options, cancellationToken);
            }
            return SaveFileAsyncCore(FilePath, options, cancellationToken);
        }

        /// <summary>
        /// Asynchronously saves the document to the specified file.
        /// </summary>
        /// <param name="filePath">The path to save the document to.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(string filePath, CancellationToken cancellationToken = default) {
            EnsureExplicitFilePath(filePath);
            return SaveFileAsyncCore(filePath, options: null, cancellationToken);
        }

        /// <summary>
        /// Asynchronously saves the document to the specified file with optional save policy settings.
        /// </summary>
        /// <param name="filePath">The path to save the document to.</param>
        /// <param name="options">Optional save policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(string filePath, WordSaveOptions? options, CancellationToken cancellationToken = default) {
            EnsureExplicitFilePath(filePath);
            return SaveFileAsyncCore(filePath, options, cancellationToken);
        }

        private static void EnsureExplicitFilePath(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            }
        }

        /// <summary>
        /// Save the WordDocument to Stream
        /// </summary>
        /// <param name="outputStream"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void Save(Stream outputStream) {
            Save(outputStream, WordFileFormat.Docx, options: null);
        }

        /// <summary>
        /// Save the WordDocument to Stream with optional save behavior.
        /// </summary>
        /// <param name="outputStream">Writable stream that receives the document content. This one-time save does not change the associated destination.</param>
        /// <param name="options">Optional save behaviors, including stream physical format selection and signed-document policy.</param>
        /// <exception cref="InvalidOperationException"></exception>
        public void Save(Stream outputStream, WordSaveOptions? options) {
            Save(outputStream, WordFileFormat.Docx, options);
        }

        /// <summary>Saves the document to a stream in the explicitly selected physical format.</summary>
        /// <param name="outputStream">Writable stream that receives the document content.</param>
        /// <param name="format">Physical DOCX or DOC format.</param>
        /// <param name="options">Optional save policy settings.</param>
        public void Save(Stream outputStream, WordFileFormat format, WordSaveOptions? options = null) {
            if (outputStream == null) {
                throw new ArgumentNullException(nameof(outputStream));
            }

            if (!outputStream.CanWrite) {
                throw new ArgumentException("Destination stream must be writable.", nameof(outputStream));
            }

            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }
            EnsureSignedDocumentSaveAllowed(options, "Save");
            PreSaving();

            if (TrySaveNativeLegacyDocToStream(outputStream, format, options)) {
                return;
            }

            EnsureLegacyDocSaveDoesNotDropImportedContent(options);
            byte[] packageBytes = CreateOpenXmlBytesAfterSave();
            OfficeStreamWriter.WriteAllBytes(outputStream, packageBytes);
        }

        /// <summary>Asynchronously saves the document to a stream as DOCX.</summary>
        /// <param name="outputStream">Writable destination stream.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(Stream outputStream, CancellationToken cancellationToken = default) {
            return SaveAsync(outputStream, WordFileFormat.Docx, options: null, cancellationToken);
        }

        /// <summary>Asynchronously saves the document to a stream as DOCX.</summary>
        /// <param name="outputStream">Writable destination stream.</param>
        /// <param name="options">Optional save policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(Stream outputStream, WordSaveOptions? options, CancellationToken cancellationToken = default) {
            return SaveAsync(outputStream, WordFileFormat.Docx, options, cancellationToken);
        }

        /// <summary>Asynchronously saves the document to a stream in the selected physical format.</summary>
        /// <param name="outputStream">Writable destination stream.</param>
        /// <param name="format">Physical DOCX or DOC format.</param>
        /// <param name="options">Optional save policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public async Task SaveAsync(
            Stream outputStream,
            WordFileFormat format,
            WordSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));
            if (!outputStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(outputStream));
            if (FileOpenAccess == FileAccess.Read) throw new InvalidOperationException("Document is read only, and cannot be saved.");

            EnsureSignedDocumentSaveAllowed(options, "SaveAsync");
            PreSaving();
            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes;
            if (format == WordFileFormat.Doc) {
                bytes = LegacyDocWriter.WriteDocument(this, options);
            } else {
                EnsureLegacyDocSaveDoesNotDropImportedContent(options);
                _wordprocessingDocument.Save();
                bytes = CreateOpenXmlBytesAfterSave();
            }

            await OfficeStreamWriter.WriteAllBytesAsync(outputStream, bytes, cancellationToken).ConfigureAwait(false);
        }

        private static bool IsLegacyDocPath(string? filePath) {
            string? extension = Path.GetExtension(filePath);
            return string.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".dot", StringComparison.OrdinalIgnoreCase);
        }

        private void SaveLegacyDocFile(string filePath, bool updateFilePath = true, WordSaveOptions? options = null) {
            byte[] legacyDocBytes = CreateLegacyDocBytesAfterPreflight(filePath, options);
            OfficeFileCommit.WriteAllBytes(filePath, legacyDocBytes);

            if (updateFilePath) {
                FilePath = filePath;
            }
        }

        private void SaveOpenXmlFile(string filePath, bool updateFilePath, WordSaveOptions? options) {
            EnsureDestinationFileWritable(filePath);

            EnsureLegacyDocSaveDoesNotDropImportedContent(options);
            _wordprocessingDocument.Save();
            OfficeFileCommit.Write(filePath, stream => {
                using (var clone = _wordprocessingDocument.Clone(stream)) {
                    AlignDocumentTypeWithFilePath(clone, filePath);
                    CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
                }

                stream.Seek(0, SeekOrigin.Begin);
                WordPackageCompatibility.NormalizeOpenOfficeRelationships(stream);
            });

            if (updateFilePath) {
                FilePath = filePath;
            }
        }

        private byte[] CreateOpenXmlBytesAfterSave(string? filePath = null) {
            using var memoryStream = new MemoryStream();
            using (var clone = _wordprocessingDocument.Clone(memoryStream, true)) {
                if (!string.IsNullOrEmpty(filePath)) {
                    AlignDocumentTypeWithFilePath(clone, filePath!);
                }
                CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
            }

            memoryStream.Seek(0, SeekOrigin.Begin);
            WordPackageCompatibility.NormalizeOpenOfficeRelationships(memoryStream);
            return memoryStream.ToArray();
        }

        private bool TrySaveNativeLegacyDocToStream(Stream destination, WordFileFormat format, WordSaveOptions? options) {
            if (format != WordFileFormat.Doc) {
                return false;
            }

            byte[] legacyDocBytes = LegacyDocWriter.WriteDocument(this, options);
            OfficeStreamWriter.WriteAllBytes(destination, legacyDocBytes);

            return true;
        }

        private static void EnsureDestinationFileWritable(string filePath) {
            if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
            }
        }

        private byte[] CreateLegacyDocBytesAfterPreflight(string filePath, WordSaveOptions? options) {
            EnsureDestinationFileWritable(filePath);

            return LegacyDocWriter.WriteDocument(
                this,
                options,
                isTemplate: string.Equals(Path.GetExtension(filePath), ".dot", StringComparison.OrdinalIgnoreCase));
        }

        private byte[] CreatePathBytesAfterPreflight(string filePath, WordSaveOptions? options) {
            if (IsLegacyDocPath(filePath)) {
                return CreateLegacyDocBytesAfterPreflight(filePath, options);
            }

            EnsureDestinationFileWritable(filePath);
            EnsureLegacyDocSaveDoesNotDropImportedContent(options);
            _wordprocessingDocument.Save();
            return CreateOpenXmlBytesAfterSave(filePath);
        }

        private void EnsureSignedDocumentSaveAllowed(WordSaveOptions? options, string operation) {
            WordSignatureInfo signatureInfo = InspectSignatures();
            if (!signatureInfo.HasSignatures) {
                return;
            }

            if (options?.SignedDocumentPolicy == WordSignedDocumentSavePolicy.AllowSignatureInvalidation) {
                return;
            }

            throw new WordSignatureSavePolicyException(operation, signatureInfo);
        }

    }
}
