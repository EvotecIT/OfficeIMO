using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint.Fluent;
using OfficeIMO.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a PowerPoint presentation providing basic create, open and save operations.
    /// </summary>
    public sealed partial class PowerPointPresentation : IDisposable {
        private PresentationDocument? _document;
        private PresentationPart _presentationPart;
        private readonly List<PowerPointSlide> _slides = new();
        private readonly string _filePath;
        private Stream? _packageStream;
        private Stream? _sourceStream;
        private bool _copyPackageToSourceOnDispose;
        private bool _saveOnDispose;
        private bool _leaveSourceStreamOpen = true;
        private PowerPointSlideSize? _slideSize;
        private bool _initialSlideUntouched = false;
        private bool _disposed = false;
        private const int StreamBufferSize = 4096;
        private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private const string SectionListUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
        private const string DefaultSectionName = "Section 1";
        private const string TableStylesResourceName = "OfficeIMO.PowerPoint.Resources.tableStyles.xml";
        private static readonly MethodInfo AddNewPartWithContentTypeMethod =
            typeof(OpenXmlPartContainer)
                .GetMethods()
                .Single(m => m.Name == "AddNewPart" &&
                             m.IsGenericMethodDefinition &&
                             m.GetParameters().Length == 2);
        private static readonly MethodInfo AddPartWithIdMethod =
            typeof(OpenXmlPartContainer)
                .GetMethods()
                .Single(m => m.Name == "AddPart" &&
                             m.IsGenericMethodDefinition &&
                             m.GetParameters().Length == 2);

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

                // After initialization, we have one slide created by PowerPointUtils
                // Track it and mark it as untouched
                if (PresentationRoot.SlideIdList != null) {
                    foreach (SlideId slideId in PresentationRoot.SlideIdList.Elements<SlideId>()) {
                        string? relId = PowerPointUtils.GetRelationshipIdValue(slideId);
                        if (!string.IsNullOrEmpty(relId)) {
                            SlidePart slidePart = (SlidePart)_presentationPart.GetPartById(relId!);
                            _slides.Add(new PowerPointSlide(slidePart));
                        }
                    }
                }
                _initialSlideUntouched = isNewPresentation && _slides.Count == 1;
            } else {
                // Loading existing presentation
                LoadExistingSlides();
                _initialSlideUntouched = false; // Existing files don't have untouched initial slide
            }

            PowerPointChartAxisIdGenerator.Initialize(_presentationPart);
        }

        private void ConfigureStreamCopy(Stream? packageStream, Stream? sourceStream, bool copyPackageToSourceOnDispose, bool leaveSourceStreamOpen) {
            _packageStream = copyPackageToSourceOnDispose ? packageStream : null;
            _sourceStream = copyPackageToSourceOnDispose ? sourceStream : null;
            _copyPackageToSourceOnDispose = copyPackageToSourceOnDispose && sourceStream != null;
            _leaveSourceStreamOpen = leaveSourceStreamOpen;
        }

        private static byte[] ReadAllBytes(Stream stream) {
            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
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

        /// <summary>
        ///     Creates a fluent wrapper for this presentation.
        /// </summary>
        public PowerPointFluentPresentation AsFluent() {
            ThrowIfDisposed();
            return new PowerPointFluentPresentation(this);
        }

        private void ThrowIfDisposed() {
            if (_disposed || _document == null) {
                throw new ObjectDisposedException(nameof(PowerPointPresentation));
            }
        }
    }
}
