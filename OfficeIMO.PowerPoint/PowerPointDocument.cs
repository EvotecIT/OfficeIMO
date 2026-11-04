using System;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a PowerPoint presentation providing create, load, and save operations.
    /// </summary>
    public sealed partial class PowerPointDocument : IDisposable, IAsyncDisposable {
        private PresentationDocument? _document;
        private PresentationPart _presentationPart;
        private readonly List<PowerPointSlide> _slides = new();
        private string _filePath;
        private Stream? _packageStream;
        private Stream? _sourceStream;
        private DocumentPersistenceMode _persistenceMode = DocumentPersistenceMode.Explicit;
        private long _maximumSlideXmlCharacters = PowerPointXmlReader.MaximumPackageXmlCharacters;
        private bool _discardChangesOnDispose;
        private string? _signedPackageOpenFingerprint;
        private PowerPointSlideSize? _slideSize;
        private bool _disposed = false;
        private const int StreamBufferSize = 4096;
        private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private const string SectionListUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
        private const string DefaultSectionName = "Section 1";
        private const string TableStylesResourceName = "OfficeIMO.PowerPoint.Resources.tableStyles.xml";
        private PowerPointDocument(PresentationDocument document, string filePath, bool isNewDocument) {
            _document = document;
            _filePath = filePath;
            _presentationPart = document.PresentationPart ?? document.AddPresentationPart();
            Properties = new PowerPointBuiltinDocumentProperties(document);
            ApplicationProperties = new PowerPointApplicationProperties(document);
            if (isNewDocument || _presentationPart.Presentation == null) {
                // New presentation - create with required initial structure
                PresentationRoot = new Presentation();
                InitializeDefaultParts();

                // InitializeDefaultParts creates the native master/layout scaffolding and one temporary slide.
                if (PresentationRoot.SlideIdList != null) {
                    foreach (SlideId slideId in PresentationRoot.SlideIdList.Elements<SlideId>()) {
                        string? relId = PowerPointUtils.GetRelationshipIdValue(slideId);
                        if (!string.IsNullOrEmpty(relId)) {
                            SlidePart slidePart = (SlidePart)_presentationPart.GetPartById(relId!);
                            _slides.Add(new PowerPointSlide(slidePart, this));
                        }
                    }
                }
                if (isNewDocument) {
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
        public PowerPointBuiltinDocumentProperties Properties { get; }

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
        ///     Replaces literal text on every slide.
        /// </summary>
        /// <param name="text">Text to find.</param>
        /// <param name="replacement">Replacement text; <c>null</c> removes each match.</param>
        /// <param name="options">Comparison settings and scope. <c>null</c> matches ordinally, ignoring case, inside or across words, in textboxes and tables but not notes.</param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>Matching follows <see cref="PowerPointSlide.ReplaceAll"/>: each textbox or table match lies within one text run.</remarks>
        public int ReplaceAll(string text, string replacement, PowerPointTextSearchOptions? options = null) {
            ThrowIfDisposed();
            OfficeTextSearch.ValidateSearchText(text, nameof(text));

            int count = 0;
            foreach (PowerPointSlide slide in Slides) {
                count += slide.ReplaceAll(text, replacement, options);
            }
            return count;
        }

        private void ThrowIfDisposed() {
            if (_disposed || _document == null) {
                throw new ObjectDisposedException(nameof(PowerPointDocument));
            }
        }
    }
}
