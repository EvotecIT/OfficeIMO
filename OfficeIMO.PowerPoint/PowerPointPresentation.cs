using System;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a PowerPoint presentation providing create, load, and save operations.
    /// </summary>
    public sealed partial class PowerPointPresentation : IDisposable, IAsyncDisposable {
        private PresentationDocument? _document;
        private PresentationPart _presentationPart;
        private readonly List<PowerPointSlide> _slides = new();
        private string _filePath;
        private Stream? _packageStream;
        private Stream? _sourceStream;
        private DocumentPersistenceMode _persistenceMode = DocumentPersistenceMode.Explicit;
        private bool _discardChangesOnDispose;
        private string? _signedPackageOpenFingerprint;
        private PowerPointSlideSize? _slideSize;
        private bool _disposed = false;
        private const int StreamBufferSize = 4096;
        private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private const string SectionListUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
        private const string DefaultSectionName = "Section 1";
        private const string TableStylesResourceName = "OfficeIMO.PowerPoint.Resources.tableStyles.xml";
        private PowerPointPresentation(PresentationDocument document, string filePath, bool isNewPresentation) {
            _document = document;
            _filePath = filePath;
            _presentationPart = document.PresentationPart ?? document.AddPresentationPart();
            BuiltinDocumentProperties = new PowerPointBuiltinDocumentProperties(document);
            ApplicationProperties = new PowerPointApplicationProperties(document);
            if (isNewPresentation || _presentationPart.Presentation == null) {
                // New presentation - create with required initial structure
                PresentationRoot = new Presentation();
                InitializeDefaultParts();

                // InitializeDefaultParts creates the native master/layout scaffolding and one temporary slide.
                if (PresentationRoot.SlideIdList != null) {
                    foreach (SlideId slideId in PresentationRoot.SlideIdList.Elements<SlideId>()) {
                        string? relId = PowerPointUtils.GetRelationshipIdValue(slideId);
                        if (!string.IsNullOrEmpty(relId)) {
                            SlidePart slidePart = (SlidePart)_presentationPart.GetPartById(relId!);
                            _slides.Add(new PowerPointSlide(slidePart));
                        }
                    }
                }
                if (isNewPresentation) {
                    ClearSlides();
                }
            } else {
                // Loading existing presentation
                LoadExistingSlides();
            }

            PowerPointChartAxisIdGenerator.Initialize(_presentationPart);
        }

        private static byte[] ReadAllBytes(Stream stream) {
            return OfficeStreamReader.ReadAllBytes(stream);
        }

        /// <summary>Gets the destination path associated with the presentation, if any.</summary>
        public string? FilePath => string.IsNullOrEmpty(_filePath) ? null : _filePath;

        /// <summary>Gets the underlying Open XML package for advanced integration scenarios.</summary>
        public PresentationDocument OpenXmlDocument {
            get {
                ThrowIfDisposed();
                return _document!;
            }
        }

        /// <summary>Gets the configured persistence behavior.</summary>
        public DocumentPersistenceMode PersistenceMode => _persistenceMode;

        /// <summary>Gets whether the presentation is editable or read-only.</summary>
        public DocumentAccessMode AccessMode {
            get {
                ThrowIfDisposed();
                return _document!.FileOpenAccess == FileAccess.Read
                    ? DocumentAccessMode.ReadOnly
                    : DocumentAccessMode.ReadWrite;
            }
        }

        /// <summary>
        ///     Collection of slides in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointSlide> Slides {
            get {
                ThrowIfDisposed();
                return _slides;
            }
        }

        /// <summary>
        ///     Built-in package properties for the presentation.
        /// </summary>
        public PowerPointBuiltinDocumentProperties BuiltinDocumentProperties { get; }

        /// <summary>
        ///     Extended application properties for the presentation.
        /// </summary>
        public PowerPointApplicationProperties ApplicationProperties { get; }

        /// <summary>
        ///     Slide size information for the presentation.
        /// </summary>
        public PowerPointSlideSize SlideSize {
            get {
                ThrowIfDisposed();
                return _slideSize ??= new PowerPointSlideSize(_presentationPart);
            }
        }

        /// <summary>
        ///     Replaces text across all slides.
        /// </summary>
        public int ReplaceText(string oldValue, string newValue, bool includeTables = true, bool includeNotes = false) {
            ThrowIfDisposed();
            if (oldValue == null) {
                throw new ArgumentNullException(nameof(oldValue));
            }
            if (oldValue.Length == 0) {
                throw new ArgumentException("Old value cannot be empty.", nameof(oldValue));
            }

            int count = 0;
            foreach (PowerPointSlide slide in Slides) {
                count += slide.ReplaceText(oldValue, newValue, includeTables, includeNotes);
            }
            return count;
        }

        private void ThrowIfDisposed() {
            if (_disposed || _document == null) {
                throw new ObjectDisposedException(nameof(PowerPointPresentation));
            }
        }
    }
}
