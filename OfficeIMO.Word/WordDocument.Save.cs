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
                    // Save current state to the memory stream
                    this._wordprocessingDocument.Save();

                    if (string.IsNullOrEmpty(filePath)) {
                        filePath = this.FilePath;
                    }

                    if (string.IsNullOrEmpty(filePath)) {
                        // No destination specified, nothing to save
                        return;
                    }

                    if (IsLegacyDocPath(filePath)) {
                        SaveLegacyDocFile(filePath);
                        if (openWord) {
                            this.Open(filePath, true);
                        }

                        return;
                    }

                    if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                        throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
                    }

                    // Allow concurrent readers (other tests may have opened the sample file with Read/ReadWrite sharing)
                    using var fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete);
                    using (var clone = this._wordprocessingDocument.Clone(fs)) {
                        AlignDocumentTypeWithFilePath(clone, filePath);
                        CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
                    }
                    fs.Seek(0, SeekOrigin.Begin);
                    Helpers.MakeOpenOfficeCompatible(fs);
                    fs.Flush();
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
            if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
            }

            byte[] packageBytes = SaveAsByteArray(saveOptions);
            byte[] encryptedBytes = OfficeEncryption.EncryptPackage(packageBytes, password);
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)) {
                fs.Write(encryptedBytes, 0, encryptedBytes.Length);
                fs.Flush();
            }
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
            byte[] packageBytes = SaveAsByteArray(saveOptions);
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
            this.Save(filePath, false, options);
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

        /// <summary>
        /// Save the document to a new file without modifying <see cref="FilePath"/> on this instance.
        /// </summary>
        /// <param name="filePath">Destination path for the cloned document.</param>
        /// <param name="openWord">Whether to open Microsoft Word after saving.</param>
        /// <returns>A new <see cref="WordDocument"/> loaded from <paramref name="filePath"/>.</returns>
        public WordDocument SaveAs(string filePath, bool openWord = false) {
            return SaveAs(filePath, openWord, options: null);
        }

        /// <summary>
        /// Save the document to a new file without modifying <see cref="FilePath"/> on this instance.
        /// </summary>
        /// <param name="filePath">Destination path for the cloned document.</param>
        /// <param name="openWord">Whether to open Microsoft Word after saving.</param>
        /// <param name="options">Optional save policy settings.</param>
        /// <returns>A new <see cref="WordDocument"/> loaded from <paramref name="filePath"/>.</returns>
        public WordDocument SaveAs(string filePath, bool openWord, WordSaveOptions? options) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }
            if (string.IsNullOrEmpty(filePath)) {
                throw new ArgumentException("File path cannot be empty", nameof(filePath));
            }

            EnsureSignedDocumentSaveAllowed(options, "SaveAs");
            PreSaving();

            if (_wordprocessingDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            try {
                if (IsLegacyDocPath(filePath)) {
                    SaveLegacyDocFile(filePath, updateFilePath: false);
                    if (openWord) {
                        Open(filePath, true);
                    }

                    WordDocument savedDocument = WordDocument.Load(filePath);
                    savedDocument.FilePath = filePath;
                    return savedDocument;
                }

                _wordprocessingDocument.Save();

                if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                    throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
                }

                using var fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete);
                using (var clone = _wordprocessingDocument.Clone(fs)) {
                    AlignDocumentTypeWithFilePath(clone, filePath);
                    CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
                }
                fs.Seek(0, SeekOrigin.Begin);
                Helpers.MakeOpenOfficeCompatible(fs);
                fs.Flush();
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                throw;
            }

            if (openWord) {
                Open(filePath, true);
            }

            return WordDocument.Load(filePath);
        }

        /// <summary>
        /// Save the document to a memory stream and return the stream's byte array.
        /// </summary>
        /// <returns>A byte array representing the saved Word document.</returns>
        public byte[] SaveAsByteArray() {
            return SaveAsByteArray(options: null);
        }

        /// <summary>
        /// Save the document to a memory stream and return the stream's byte array.
        /// </summary>
        /// <param name="options">Optional save policy settings.</param>
        /// <returns>A byte array representing the saved Word document.</returns>
        public byte[] SaveAsByteArray(WordSaveOptions? options) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }

            EnsureSignedDocumentSaveAllowed(options, "SaveAsByteArray");
            PreSaving();

            if (_wordprocessingDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            try {
                _wordprocessingDocument.Save();

                using var memoryStream = new MemoryStream();
                using (var clone = _wordprocessingDocument.Clone(memoryStream, true)) {
                    CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
                }

                memoryStream.Seek(0, SeekOrigin.Begin);
                Helpers.MakeOpenOfficeCompatible(memoryStream);
                memoryStream.Flush();

                return memoryStream.ToArray();
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                throw;
            }
        }

        /// <summary>
        /// Save the document to a new <see cref="MemoryStream"/>.
        /// </summary>
        /// <returns>A memory stream containing the saved document.</returns>
        public MemoryStream SaveAsMemoryStream() {
            return SaveAsMemoryStream(options: null);
        }

        /// <summary>
        /// Save the document to a new <see cref="MemoryStream"/>.
        /// </summary>
        /// <param name="options">Optional save policy settings.</param>
        /// <returns>A memory stream containing the saved document.</returns>
        public MemoryStream SaveAsMemoryStream(WordSaveOptions? options) {
            var stream = new MemoryStream();
            Save(stream, options);
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
        }

        /// <summary>
        /// Clone the document to the specified stream and return a new instance loaded from it.
        /// </summary>
        /// <param name="outputStream">Target stream that must support reading and seeking.</param>
        /// <returns>A new <see cref="WordDocument"/> loaded from <paramref name="outputStream"/>.</returns>
        public WordDocument SaveAs(Stream outputStream) {
            return SaveAs(outputStream, options: null);
        }

        /// <summary>
        /// Clone the document to the specified stream and return a new instance loaded from it.
        /// </summary>
        /// <param name="outputStream">Target stream that must support reading and seeking.</param>
        /// <param name="options">Optional save policy settings.</param>
        /// <returns>A new <see cref="WordDocument"/> loaded from <paramref name="outputStream"/>.</returns>
        public WordDocument SaveAs(Stream outputStream, WordSaveOptions? options) {
            if (outputStream == null) {
                throw new ArgumentNullException(nameof(outputStream));
            }
            if (!outputStream.CanSeek) {
                throw new ArgumentException("Stream must support seeking", nameof(outputStream));
            }

            Save(outputStream, options);
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
            EnsureSignedDocumentSaveAllowed(options, "SaveAsync");
            PreSaving();

            if (this._wordprocessingDocument != null) {
                try {
                    this._wordprocessingDocument.Save();

                    if (string.IsNullOrEmpty(filePath)) {
                        filePath = this.FilePath;
                    }

                    if (string.IsNullOrEmpty(filePath)) {
                        return;
                    }

                    if (IsLegacyDocPath(filePath)) {
                        byte[] legacyDocBytes = CreateLegacyDocBytesAfterPreflight(filePath);
                        using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete, 4096, FileOptions.Asynchronous)) {
                            await fs.WriteAsync(legacyDocBytes, 0, legacyDocBytes.Length, cancellationToken);
                            await fs.FlushAsync(cancellationToken);
                        }
                        FilePath = filePath;
                        if (openWord) {
                            this.Open(filePath, true);
                        }

                        return;
                    }

                    if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                        throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
                    }

                    var directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory)) {
                        var dirInfo = new DirectoryInfo(directory);
                        if (dirInfo.Attributes.HasFlag(FileAttributes.ReadOnly)) {
                            throw new IOException($"Failed to save to '{filePath}'. The directory is read-only.");
                        }
                    }

                    using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete, 4096, FileOptions.Asynchronous)) {
                        using (var clone = this._wordprocessingDocument.Clone(fs)) {
                            AlignDocumentTypeWithFilePath(clone, filePath);
                            CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
                        }
                        fs.Seek(0, SeekOrigin.Begin);
                        Helpers.MakeOpenOfficeCompatible(fs);
                        await fs.FlushAsync(cancellationToken);
                    }
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
            Save(outputStream, options: null);
        }

        /// <summary>
        /// Save the WordDocument to Stream with optional save behavior.
        /// </summary>
        /// <param name="outputStream">Writable stream that receives the document content.</param>
        /// <param name="options">Optional save behaviors, including stream physical format selection and signed-document policy.</param>
        /// <exception cref="InvalidOperationException"></exception>
        public void Save(Stream outputStream, WordSaveOptions? options) {
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

            if (TrySaveNativeLegacyDocToStream(outputStream, options)) {
                return;
            }

            // Clone document once and copy package properties in the same operation
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

        private static bool IsLegacyDocPath(string? filePath) {
            return string.Equals(Path.GetExtension(filePath), ".doc", StringComparison.OrdinalIgnoreCase);
        }

        private void SaveLegacyDocFile(string filePath, bool updateFilePath = true) {
            byte[] legacyDocBytes = CreateLegacyDocBytesAfterPreflight(filePath);
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete)) {
                fs.Write(legacyDocBytes, 0, legacyDocBytes.Length);
                fs.Flush();
            }

            if (updateFilePath) {
                FilePath = filePath;
            }
        }

        private bool TrySaveNativeLegacyDocToStream(Stream destination, WordSaveOptions? options) {
            if (options?.StreamFormat != WordStreamSaveFormat.LegacyDoc) {
                return false;
            }

            byte[] legacyDocBytes = LegacyDocWriter.WriteDocument(this);
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

        private byte[] CreateLegacyDocBytesAfterPreflight(string filePath) {
            if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
            }

            return LegacyDocWriter.WriteDocument(this);
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
