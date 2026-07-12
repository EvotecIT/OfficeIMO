using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Core;
using OfficeIMO.Core.Internal;
using OfficeIMO.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
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
                Save(_sourceStream);
                return;
            }
            throw new InvalidOperationException(
                "The presentation has no associated destination. Use Save(string) or Save(Stream).");
        }

        /// <summary>Saves the presentation to a file and associates that path with subsequent <see cref="Save()"/> calls.</summary>
        public void Save(string filePath) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            byte[] packageBytes = CreatePackageBytesForSave();
            OfficeFileCommit.WriteAllBytes(filePath, packageBytes);
            _filePath = filePath;
            _discardChangesOnDispose = false;
        }

        /// <summary>
        ///     Saves the presentation to the provided stream.
        /// </summary>
        public void Save(Stream destination) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            OfficeStreamWriter.WriteAllBytes(destination, CreatePackageBytesForSave());
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
            using PowerPointPresentation exported = Create(destination);
            exported.ImportSlide(this, slideIndex);
            exported.Save(destination);
        }

        /// <summary>
        ///     Saves the presentation as a password-encrypted Office Open XML package.
        /// </summary>
        /// <param name="filePath">Destination path for the encrypted presentation.</param>
        /// <param name="password">Password used to encrypt the presentation package.</param>
        /// <param name="openPowerPoint">Whether to open the saved file after writing.</param>
        public void SaveEncrypted(string filePath, string password, bool openPowerPoint = false) {
            ThrowIfDisposed();
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (filePath.Length == 0) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
            }

            using var packageStream = new MemoryStream();
            Save(packageStream);
            byte[] encryptedBytes = OfficeEncryption.EncryptPackage(packageStream.ToArray(), password);
            OfficeFileCommit.WriteAllBytes(filePath, encryptedBytes);

            if (openPowerPoint) {
                Helpers.Open(filePath, true);
            }
        }

        /// <summary>
        ///     Saves the presentation as a password-encrypted Office Open XML package to a stream.
        /// </summary>
        /// <param name="destination">Writable stream receiving the encrypted presentation.</param>
        /// <param name="password">Password used to encrypt the presentation package.</param>
        public void SaveEncrypted(Stream destination, string password) {
            ThrowIfDisposed();
            if (password == null) throw new ArgumentNullException(nameof(password));

            using var packageStream = new MemoryStream();
            Save(packageStream);
            byte[] encryptedBytes = OfficeEncryption.EncryptPackage(packageStream.ToArray(), password);
            OfficeStreamWriter.WriteAllBytes(destination, encryptedBytes);
        }

    }
}
