using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Shared;
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

        /// <summary>
        /// Open WordDocument in Microsoft Word (if Word is present)
        /// </summary>
        /// <param name="openWord"></param>
        public void Open(bool openWord = true) {
            this.Open("", openWord);
        }

        /// <summary>
        /// Open WordDocument in Microsoft Word (if Word is present)
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="openWord"></param>
        public void Open(string filePath = "", bool openWord = true) {
            if (filePath == "") {
                filePath = this.FilePath;
            }

            Helpers.Open(filePath, openWord);
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

        /// <summary>
        /// Save WordDocument to filePath (SaveAs), and open the file in Microsoft Word
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="openWord"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void Save(string filePath, bool openWord) {
            Save(filePath, openWord, options: null);
        }

        /// <summary>
        /// Save WordDocument to filePath (SaveAs), optionally opening the file in Microsoft Word.
        /// </summary>
        /// <param name="filePath">Destination path. When empty, uses the current <see cref="FilePath"/>.</param>
        /// <param name="openWord">Whether to open Microsoft Word after saving.</param>
        /// <param name="options">Optional save policy settings.</param>
        public void Save(string filePath, bool openWord, WordSaveOptions? options) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }
            EnsureSignedDocumentSaveAllowed(options, "Save");
            PreSaving();

            if (this._wordprocessingDocument != null) {
                try {
                    if (string.IsNullOrEmpty(filePath)) {
                        filePath = this.FilePath;
                    }

                    if (string.IsNullOrEmpty(filePath)) {
                        throw new InvalidOperationException("This document is not associated with a file path. Provide a file path or save to a writable stream.");
                    }

                    if (IsLegacyDocPath(filePath)) {
                        SaveLegacyDocFile(filePath, options: options);
                        if (openWord) {
                            this.Open(filePath, true);
                        }

                        return;
                    }

                    SaveOpenXmlFile(filePath, updateFilePath: true, options);
                } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                    throw;
                }
            } else {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            if (openWord) {
                this.Open(filePath, true);
            }
        }

        /// <summary>
        /// Saves the document as a password-encrypted Office Open XML package.
        /// </summary>
        /// <param name="filePath">Destination path. When empty, uses the current <see cref="FilePath"/>.</param>
        /// <param name="password">Password used to encrypt the document package.</param>
        /// <param name="openWord">Whether to open the saved file after writing.</param>
        public void SaveEncrypted(string filePath, string password, bool openWord = false) {
            SaveEncrypted(filePath, password, openWord, saveOptions: null);
        }

        /// <summary>
        /// Saves the document as a password-encrypted Office Open XML package.
        /// </summary>
        /// <param name="filePath">Destination path. When empty, uses the current <see cref="FilePath"/>.</param>
        /// <param name="password">Password used to encrypt the document package.</param>
        /// <param name="openWord">Whether to open the saved file after writing.</param>
        /// <param name="saveOptions">Optional save policy settings.</param>
        public void SaveEncrypted(string filePath, string password, bool openWord, WordSaveOptions? saveOptions) {
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (string.IsNullOrEmpty(filePath)) {
                filePath = FilePath;
            }
            if (string.IsNullOrEmpty(filePath)) {
                throw new InvalidOperationException("This document is not associated with a file path. Provide a file path or call SaveEncrypted(Stream, ...).");
            }
            EnsureDestinationFileWritable(filePath);

            byte[] packageBytes = ToDocx(saveOptions);
            byte[] encryptedBytes = OfficeEncryption.EncryptPackage(packageBytes, password);
            OfficeFileCommit.WriteAllBytes(filePath, encryptedBytes);
            FilePath = filePath;

            if (openWord) {
                Open(filePath, true);
            }
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
            byte[] packageBytes = ToDocx(saveOptions);
            OfficeEncryption.EncryptPackageToStream(packageBytes, password, destination);
        }

        /// <summary>
        /// Save WordDocument to where it was open from
        /// </summary>
        public void Save() {
            this.Save(false);
        }

        /// <summary>
        /// Save WordDocument to given filePath
        /// </summary>
        /// <param name="filePath"></param>
        public void Save(string filePath) {
            this.Save(filePath, false);
        }

        /// <summary>
        /// Save WordDocument to the given file path with optional save policy settings.
        /// </summary>
        /// <param name="filePath">Destination path.</param>
        /// <param name="options">Optional save policy settings.</param>
        public void Save(string filePath, WordSaveOptions? options) {
            this.Save(filePath, options?.OpenAfterSave == true, options);
        }

        /// <summary>
        /// Save WordDocument and open it in Microsoft Word (if Word is present)
        /// </summary>
        /// <param name="openWord"></param>
        public void Save(bool openWord) {
            Save(openWord, options: null);
        }

        /// <summary>
        /// Save WordDocument and optionally open it in Microsoft Word.
        /// </summary>
        /// <param name="openWord">Whether to open Microsoft Word after saving.</param>
        /// <param name="options">Optional save policy settings.</param>
        public void Save(bool openWord, WordSaveOptions? options) {
            if (string.IsNullOrEmpty(this.FilePath) && this.OriginalStream != null) {
                this.Save(this.OriginalStream, options);
            } else {
                this.Save("", openWord, options);
            }
        }

        // Note: Save() already normalizes table grids for consistent viewing across
        // Word Online/Google Docs without changing authoring semantics. No extra
        // save variants are needed.

        private WordDocument SaveCopyCore(string filePath, bool openWord, WordSaveOptions? options) {
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

            try {
                if (IsLegacyDocPath(filePath)) {
                    SaveLegacyDocFile(filePath, updateFilePath: false, options);
                    if (openWord) {
                        Open(filePath, true);
                    }

                    WordDocument savedDocument = WordDocument.Load(filePath);
                    savedDocument.FilePath = filePath;
                    return savedDocument;
                }

                SaveOpenXmlFile(filePath, updateFilePath: false, options);
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                throw;
            }

            if (openWord) {
                Open(filePath, true);
            }

            return WordDocument.Load(filePath);
        }

        /// <summary>Saves an independent copy and returns the document loaded from that copy.</summary>
        /// <param name="filePath">Destination path.</param>
        /// <param name="options">Optional save settings, including <see cref="WordSaveOptions.OpenAfterSave"/>.</param>
        /// <returns>A new document associated with <paramref name="filePath"/>. This instance keeps its current path.</returns>
        public WordDocument SaveCopy(string filePath, WordSaveOptions? options = null) {
            return SaveCopyCore(filePath, options?.OpenAfterSave == true, options);
        }

        /// <summary>Encodes the document as an Office Open XML DOCX package.</summary>
        /// <param name="options">Optional save policy settings.</param>
        /// <returns>DOCX package bytes.</returns>
        public byte[] ToDocx(WordSaveOptions? options = null) {
            return ToWordBytes(WordFileFormat.Docx, options);
        }

        /// <summary>Encodes the document as a Word 97-2003 binary DOC file.</summary>
        /// <param name="options">Optional save policy settings.</param>
        /// <returns>DOC compound-file bytes.</returns>
        public byte[] ToDoc(WordSaveOptions? options = null) {
            return ToWordBytes(WordFileFormat.Doc, options);
        }

        /// <summary>Encodes the document as DOCX in a new memory stream.</summary>
        /// <param name="options">Optional save policy settings.</param>
        /// <returns>A writable memory stream positioned at the beginning.</returns>
        public MemoryStream ToDocxStream(WordSaveOptions? options = null) {
            return new MemoryStream(ToDocx(options));
        }

        /// <summary>Encodes the document as DOC in a new memory stream.</summary>
        /// <param name="options">Optional save policy settings.</param>
        /// <returns>A writable memory stream positioned at the beginning.</returns>
        public MemoryStream ToDocStream(WordSaveOptions? options = null) {
            return new MemoryStream(ToDoc(options));
        }

        private byte[] ToWordBytes(WordFileFormat format, WordSaveOptions? options) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }

            EnsureSignedDocumentSaveAllowed(options, format == WordFileFormat.Doc ? "ToDoc" : "ToDocx");
            PreSaving();

            if (_wordprocessingDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            if (format == WordFileFormat.Doc) {
                return LegacyDocWriter.WriteDocument(this, options);
            }

            EnsureLegacyDocSaveDoesNotDropImportedContent(options);

            try {
                _wordprocessingDocument.Save();
                return CreateOpenXmlBytesAfterSave();
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                throw;
            }
        }

        /// <summary>Saves an independent stream copy and returns a document loaded from it.</summary>
        /// <param name="outputStream">Readable, writable, seekable destination stream.</param>
        /// <param name="format">Physical DOCX or DOC format.</param>
        /// <param name="options">Optional save settings.</param>
        /// <returns>A new document backed by <paramref name="outputStream"/>.</returns>
        public WordDocument SaveCopy(Stream outputStream, WordFileFormat format = WordFileFormat.Docx, WordSaveOptions? options = null) {
            if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));
            if (!outputStream.CanRead || !outputStream.CanWrite || !outputStream.CanSeek) {
                throw new ArgumentException("Stream must support reading, writing, and seeking.", nameof(outputStream));
            }

            Stream originalStream = OriginalStream;
            try {
                Save(outputStream, format, options);
            } finally {
                OriginalStream = originalStream;
            }

            outputStream.Seek(0, SeekOrigin.Begin);
            return WordDocument.Load(outputStream);
        }

        /// <summary>
        /// Asynchronously saves the document.
        /// </summary>
        /// <param name="filePath">Optional path to save to.</param>
        /// <param name="openWord">Whether to open Word after saving.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public async Task SaveAsync(string filePath, bool openWord, CancellationToken cancellationToken = default) {
            await SaveAsync(filePath, openWord, options: null, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously saves the document.
        /// </summary>
        /// <param name="filePath">Optional path to save to.</param>
        /// <param name="openWord">Whether to open Word after saving.</param>
        /// <param name="options">Optional save policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public async Task SaveAsync(string filePath, bool openWord, WordSaveOptions? options, CancellationToken cancellationToken = default) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }

            if (string.IsNullOrEmpty(filePath) && string.IsNullOrEmpty(FilePath) && OriginalStream != null) {
                await SaveAsync(OriginalStream, WordFileFormat.Docx, options, cancellationToken).ConfigureAwait(false);
                return;
            }

            EnsureSignedDocumentSaveAllowed(options, "SaveAsync");
            PreSaving();

            if (this._wordprocessingDocument != null) {
                try {
                    if (string.IsNullOrEmpty(filePath)) {
                        filePath = this.FilePath;
                    }

                    if (string.IsNullOrEmpty(filePath)) {
                        throw new InvalidOperationException("This document is not associated with a file path. Provide a file path or save to a writable stream.");
                    }

                    if (IsLegacyDocPath(filePath)) {
                        byte[] legacyDocBytes = CreateLegacyDocBytesAfterPreflight(filePath, options);
                        await OfficeFileCommit.WriteAllBytesAsync(filePath, legacyDocBytes, cancellationToken: cancellationToken).ConfigureAwait(false);
                        FilePath = filePath;
                        if (openWord) {
                            this.Open(filePath, true);
                        }

                        return;
                    }

                    EnsureDestinationFileWritable(filePath);
                    this._wordprocessingDocument.Save();
                    EnsureLegacyDocSaveDoesNotDropImportedContent(options);
                    byte[] openXmlBytes = CreateOpenXmlBytesAfterSave(filePath);
                    await OfficeFileCommit.WriteAllBytesAsync(filePath, openXmlBytes, cancellationToken: cancellationToken).ConfigureAwait(false);
                    FilePath = filePath;
                } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                    throw;
                }
            } else {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            if (openWord) {
                this.Open(filePath, true);
            }
        }

        /// <summary>
        /// Asynchronously saves the document to where it was open from.
        /// </summary>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(CancellationToken cancellationToken = default) {
            return SaveAsync("", false, cancellationToken);
        }

        /// <summary>
        /// Asynchronously saves the document to the specified file.
        /// </summary>
        /// <param name="filePath">The path to save the document to.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(string filePath, CancellationToken cancellationToken = default) {
            return SaveAsync(filePath, false, cancellationToken);
        }

        /// <summary>
        /// Asynchronously saves the document to the specified file with optional save policy settings.
        /// </summary>
        /// <param name="filePath">The path to save the document to.</param>
        /// <param name="options">Optional save policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(string filePath, WordSaveOptions? options, CancellationToken cancellationToken = default) {
            return SaveAsync(filePath, false, options, cancellationToken);
        }

        /// <summary>
        /// Asynchronously saves the document and opens it in Microsoft Word (if Word is present).
        /// </summary>
        /// <param name="openWord">Whether to open Word after saving.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(bool openWord, CancellationToken cancellationToken = default) {
            return SaveAsync("", openWord, cancellationToken);
        }

        /// <summary>
        /// Asynchronously saves the document and opens it in Microsoft Word when requested.
        /// </summary>
        /// <param name="openWord">Whether to open Word after saving.</param>
        /// <param name="options">Optional save policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(bool openWord, WordSaveOptions? options, CancellationToken cancellationToken = default) {
            return SaveAsync("", openWord, options, cancellationToken);
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
        /// <param name="outputStream">Writable stream that receives the document content.</param>
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

            // Clone document once and copy package properties in the same operation
            PrepareDestinationStreamForWrite(outputStream);
            using (var clone = this._wordprocessingDocument.Clone(outputStream)) {
                CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
            }

            // Keep stream-based saves aligned with file-based saves when the destination
            // supports the read/write/seek semantics required by Package.Open.
            if (outputStream.CanRead && outputStream.CanWrite && outputStream.CanSeek) {
                outputStream.Seek(0, SeekOrigin.Begin);
                Helpers.MakeOpenOfficeCompatible(outputStream);
            }

            OriginalStream = outputStream;

            if (outputStream.CanSeek) {
                outputStream.Seek(0, SeekOrigin.Begin);
            }
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

            PrepareDestinationStreamForWrite(outputStream);
#if NET6_0_OR_GREATER
            await outputStream.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
            await outputStream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
            try { await outputStream.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
            // A pathless Save() targets OriginalStream as DOCX. Do not bind a legacy
            // DOC stream here or a later Save()/auto-save could replace it with OOXML.
            if (format == WordFileFormat.Docx) {
                OriginalStream = outputStream;
            }
            if (outputStream.CanSeek) outputStream.Seek(0, SeekOrigin.Begin);
        }

        private static bool IsLegacyDocPath(string? filePath) {
            return string.Equals(Path.GetExtension(filePath), ".doc", StringComparison.OrdinalIgnoreCase);
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
                Helpers.MakeOpenOfficeCompatible(stream);
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
            Helpers.MakeOpenOfficeCompatible(memoryStream);
            return memoryStream.ToArray();
        }

        private bool TrySaveNativeLegacyDocToStream(Stream destination, WordFileFormat format, WordSaveOptions? options) {
            if (format != WordFileFormat.Doc) {
                return false;
            }

            byte[] legacyDocBytes = LegacyDocWriter.WriteDocument(this, options);
            PrepareDestinationStreamForWrite(destination);
            destination.Write(legacyDocBytes, 0, legacyDocBytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }

            if (destination.CanSeek) {
                destination.Seek(0, SeekOrigin.Begin);
            }

            return true;
        }

        private static void PrepareDestinationStreamForWrite(Stream destination) {
            if (!destination.CanSeek) {
                return;
            }

            destination.Seek(0, SeekOrigin.Begin);
            destination.SetLength(0);
        }

        private static void EnsureDestinationFileWritable(string filePath) {
            if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
            }
        }

        private byte[] CreateLegacyDocBytesAfterPreflight(string filePath, WordSaveOptions? options) {
            EnsureDestinationFileWritable(filePath);

            return LegacyDocWriter.WriteDocument(this, options);
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
